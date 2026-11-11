import AppKit
import ApplicationServices

struct Failure: Error { let message: String; init(_ message: String) { self.message = message } }
func attr(_ e: AXUIElement, _ name: String) -> Any? {
    var value: CFTypeRef?
    guard AXUIElementCopyAttributeValue(e, name as CFString, &value) == .success else { return nil }
    return value
}
func string(_ e: AXUIElement, _ name: String) -> String { attr(e, name) as? String ?? "" }
func elements(_ e: AXUIElement, _ name: String = "AXChildren") -> [AXUIElement] { attr(e, name) as? [AXUIElement] ?? [] }
func descendants(_ root: AXUIElement, stoppingAt predicate: ((AXUIElement) -> Bool)? = nil) -> [AXUIElement] {
    var pending = [root], visited = Set<CFHashCode>(), result = [AXUIElement]()
    while let node = pending.popLast(), result.count < 3000 {
        guard visited.insert(CFHash(node)).inserted else { continue }
        if predicate?(node) == true { return [node] }
        result.append(node)
        pending.append(contentsOf: elements(node).reversed())
    }
    return predicate == nil ? result : []
}
func wait<T>(_ failure: String, _ operation: () -> T?) throws -> T {
    for _ in 0..<40 {
        if let result = operation() { return result }
        Thread.sleep(forTimeInterval: 0.25)
    }
    throw Failure(failure)
}
func set(_ e: AXUIElement, _ key: String, _ value: Any, _ failure: String) throws {
    guard AXUIElementSetAttributeValue(e, key as CFString, value as CFTypeRef) == .success else {
        throw Failure(failure)
    }
}
func action(_ e: AXUIElement, _ name: String) -> Bool {
    AXUIElementPerformAction(e, name as CFString) == .success
}
func select(_ row: AXUIElement, in table: AXUIElement) throws {
    if AXUIElementSetAttributeValue(table, "AXSelectedRows" as CFString, [row] as CFArray) == .success { return }
    try set(row, "AXSelected", true, "Outlook did not allow selecting the row.")
}

func run() throws -> [String: Any] {
    guard AXIsProcessTrusted() else { throw Failure("Grant macOS Accessibility to the app that runs this MCP.") }
    let data = FileHandle.standardInput.readDataToEndOfFile()
    guard let input = try JSONSerialization.jsonObject(with: data) as? [String: Any],
          let account = input["account"] as? String, let operation = input["operation"] as? String,
          let app = NSRunningApplication.runningApplications(withBundleIdentifier: "com.microsoft.Outlook").first else {
        throw Failure("Open Outlook for Mac and sign in to the requested account.")
    }
    let application = AXUIElementCreateApplication(app.processIdentifier)
    guard let window = elements(application, "AXWindows").first(where: { string($0, "AXIdentifier") == "OutlookWindow" }) else {
        throw Failure("Open the main Outlook mail window.")
    }
    func foreground() throws {
        if NSWorkspace.shared.frontmostApplication?.processIdentifier == app.processIdentifier { return }
        let opener = Process()
        opener.executableURL = URL(fileURLWithPath: "/usr/bin/open")
        opener.arguments = ["-b", "com.microsoft.Outlook"]
        try opener.run()
        opener.waitUntilExit()
        _ = try wait("Outlook must be in the foreground.") {
            NSWorkspace.shared.frontmostApplication?.processIdentifier == app.processIdentifier ? true : nil
        }
    }
    func key(_ code: CGKeyCode, flags: CGEventFlags = [], text: String? = nil) throws {
        guard NSWorkspace.shared.frontmostApplication?.processIdentifier == app.processIdentifier else {
            throw Failure("Outlook lost focus.")
        }
        for down in [true, false] {
            guard let event = CGEvent(keyboardEventSource: nil, virtualKey: code, keyDown: down) else {
                throw Failure("Could not create the Outlook input event.")
            }
            event.flags = flags
            if let text, down {
                let units = Array(text.utf16)
                units.withUnsafeBufferPointer {
                    event.keyboardSetUnicodeString(stringLength: units.count, unicodeString: $0.baseAddress)
                }
            }
            event.post(tap: .cghidEventTap)
        }
    }
    let inboxTitle = "Inbox • " + account
    let searchTitle = "Searching \"" + account + "\""
    func title() -> String { string(window, "AXTitle") }
    func find(_ key: String, _ value: String) -> AXUIElement? {
        descendants(window, stoppingAt: { string($0, key) == value }).first
    }
    func scopeIsCurrent() -> Bool {
        guard let scope = find("AXIdentifier", "searchScopeButton") else { return false }
        return string(scope, "AXValue") == "Current Mailbox"
    }
    func messageTable() -> AXUIElement? {
        descendants(window, stoppingAt: { string($0, "AXRole") == "AXTable" && string($0, "AXDescription") == "Message List" }).first
    }
    func messages(_ table: AXUIElement) -> [(AXUIElement, String)] {
        let rows = elements(table, "AXRows")
        return (rows.isEmpty ? elements(table) : rows).compactMap { row in
            guard string(row, "AXRole") == "AXRow",
                  let cell = elements(row).first(where: { string($0, "AXRole") == "AXCell" }) else { return nil }
            let description = string(cell, "AXDescription")
            return description.contains(",") ? (row, description) : nil
        }
    }

    if operation == "search", let query = input["query"] as? String, let limit = input["limit"] as? Int {
        guard let outline = find("AXRole", "AXOutline") else { throw Failure("Outlook mailbox sidebar is unavailable.") }
        let rows = elements(outline, "AXRows")
        let folders = rows.isEmpty ? elements(outline) : rows
        func label(_ row: AXUIElement) -> String {
            let cell = elements(row).first ?? row
            return [row, cell].flatMap { [$0] + elements($0) }.flatMap { item in
                ["AXDescription", "AXValue", "AXTitle"].map { string(item, $0) }
            }.joined(separator: "\n")
        }
        guard let header = folders.firstIndex(where: { label($0).split(separator: "\n").contains(Substring(account)) }),
              let inbox = folders.dropFirst(header + 1).prefix(3).first(where: {
                  label($0).split(separator: "\n").contains { $0 == "Inbox" || $0.hasPrefix("Inbox;") }
              }) else { throw Failure("Expand the requested account in Outlook's sidebar.") }
        try foreground()
        if title().hasPrefix("Searching ") {
            try key(53)
            _ = try wait("Outlook did not exit the previous search.") { title().hasPrefix("Searching ") ? nil : true }
        }
        if title() != inboxTitle {
            try select(inbox, in: outline)
            _ = try wait("Outlook did not switch to the requested account.") { title() == inboxTitle ? true : nil }
        }
        guard let field = find("AXIdentifier", "Search Bar") else { throw Failure("Outlook search field is unavailable.") }
        try set(field, "AXFocused", true, "Could not focus Outlook search.")
        _ = try wait("Outlook search is not limited to Current Mailbox.") { scopeIsCurrent() ? true : nil }
        guard attr(field, "AXFocused") as? Bool == true else { throw Failure("Outlook search is not focused.") }
        try key(0, flags: .maskCommand)
        try key(0, text: query)
        _ = try wait("Outlook did not accept the search query.") {
            string(field, "AXValue") == query ? true : nil
        }
        try key(36)
        _ = try wait("Outlook did not confirm the requested mailbox scope.") {
            title() == searchTitle && scopeIsCurrent() ? true : nil
        }
        let results = try wait("Outlook did not expose completed search results.") { () -> [String]? in
            guard title() == searchTitle, let table = messageTable() else { return nil }
            let descriptions = messages(table).map { $0.1 }
            if !descriptions.isEmpty { return Array(descriptions.prefix(limit)) }
            let empty = descendants(window).contains { item in
                ["AXValue", "AXTitle", "AXDescription"].contains { key in
                    let value = string(item, key)
                    return value == "No Results" || value == "No results found" || value == "No results"
                }
            }
            return empty ? [] : nil
        }
        return ["rows": results]
    }

    if operation == "read", let expected = input["row"] as? String, let limit = input["max_chars"] as? Int, let index = input["index"] as? Int,
       let query = input["query"] as? String {
        guard title() == searchTitle, scopeIsCurrent(), let table = messageTable(),
              let field = find("AXIdentifier", "Search Bar"), string(field, "AXValue") == query else {
            throw Failure("The scoped Outlook search has changed. Search again.")
        }
        let matches = messages(table)
        guard matches.indices.contains(index), matches[index].1 == expected else {
            throw Failure("The Outlook search result moved. Search again.")
        }
        let row = matches[index].0
        func selected() -> Bool {
            attr(row, "AXSelected") as? Bool == true ||
                elements(table, "AXSelectedRows").contains { CFEqual($0, row) }
        }
        func bodyText(_ area: AXUIElement) -> String {
            descendants(area).filter { string($0, "AXRole") == "AXStaticText" }.map { node in
                let value = string(node, "AXValue")
                return value.isEmpty ? string(node, "AXTitle") : value
            }.filter { !$0.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty }.joined(separator: "\n")
        }
        let previous = find("AXRole", "AXWebArea").map(bodyText)
        let wasSelected = selected()
        if !wasSelected {
            guard let position = attr(row, "AXPosition"), let size = attr(row, "AXSize") else {
                throw Failure("The Outlook result is outside the visible list.")
            }
            var origin = CGPoint.zero, bounds = CGSize.zero
            guard AXValueGetValue(position as! AXValue, .cgPoint, &origin),
                  AXValueGetValue(size as! AXValue, .cgSize, &bounds), bounds.width > 0, bounds.height > 0 else {
                throw Failure("Could not locate the Outlook result.")
            }
            let point = CGPoint(x: origin.x + bounds.width / 2, y: origin.y + bounds.height / 2)
            try foreground()
            _ = action(window, "AXRaise")
            var hit: AXUIElement?
            guard AXUIElementCopyElementAtPosition(AXUIElementCreateSystemWide(), Float(point.x), Float(point.y), &hit) == .success else {
                throw Failure("The Outlook result is not visible.")
            }
            for _ in 0..<12 {
                guard let element = hit, !CFEqual(element, row) else { break }
                var parent: CFTypeRef?
                if AXUIElementCopyAttributeValue(element, "AXParent" as CFString, &parent) != .success { hit = nil; break }
                hit = parent.map { unsafeDowncast($0, to: AXUIElement.self) }
            }
            guard let hit, CFEqual(hit, row) else { throw Failure("The Outlook result is not visible.") }
            for type in [CGEventType.leftMouseDown, .leftMouseUp] {
                CGEvent(mouseEventSource: nil, mouseType: type, mouseCursorPosition: point, mouseButton: .left)?
                    .post(tap: .cghidEventTap)
            }
        }
        let body = try wait("Outlook did not expose a matching message body.") { () -> String? in
            guard title() == searchTitle, selected(), let area = find("AXRole", "AXWebArea") else { return nil }
            let text = bodyText(area)
            if !wasSelected && text == previous { return nil }
            return text.isEmpty ? nil : text
        }
        return ["body": String(body.prefix(limit)), "truncated": body.count > limit]
    }
    throw Failure("Invalid Outlook operation.")
}

do {
    let result = try run()
    FileHandle.standardOutput.write(try JSONSerialization.data(withJSONObject: result))
} catch {
    let message = (error as? Failure)?.message ?? "Outlook accessibility request failed."
    FileHandle.standardOutput.write(try! JSONSerialization.data(withJSONObject: ["error": message]))
}
