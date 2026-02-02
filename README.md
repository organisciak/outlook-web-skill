# Outlook Web Plugin for Claude Code

A Claude Code plugin that enables browser automation of Outlook Web (outlook.office.com) for email and calendar management using playwright-cli.

## Features

- **Email Operations**
  - Read inbox and emails
  - Compose drafts (safe by default - never sends without explicit confirmation)
  - Reply to and forward emails
  - Search emails
  - Navigate folders (Drafts, Sent Items, Archive, etc.)

- **Calendar Operations**
  - View calendar (day/week/month views)
  - See upcoming events
  - Access event details

- **Session Management**
  - Persistent authentication across commands
  - Headless operation (after initial login)
  - Automatic draft saving

- **Optimized for LLMs**
  - Text extraction methods for efficient token usage (70-90% reduction)
  - Batch operations with run-code
  - Resource blocking options for faster page loads

## Prerequisites

**Required:**
1. A took taht supports skills (Claude Code, Codex, OpenClaw)
3. **playwright-cli** - Install the official Playwright CLI plugin:
   ```bash
   /plugin marketplace add anthropics/playwright-cli-plugin
   /plugin install playwright-cli
   ```

   Or install via npm:
   ```bash
   npm install -g @playwright/cli@latest
   ```

> **Note:** This plugin requires `playwright-cli` to interact with Outlook Web. Ensure it's installed and working before using this skill.

## Installation

### Via Plugin Marketplace (Recommended)

Add the marketplace and install the plugin:

```bash
/plugin marketplace add organisciak/outlook-web-plugin
/plugin install outlook-web
```

### Manual Installation

For agents without automatic plugin support, install the skill directly:

```bash
mkdir -p ~/.claude/skills/outlook-web
curl -o ~/.claude/skills/outlook-web/SKILL.md \
  https://raw.githubusercontent.com/organisciak/outlook-web-plugin/main/skills/outlook-web/SKILL.md
```

Optionally, install the reference documentation:

```bash
mkdir -p ~/.claude/skills/outlook-web/references
curl -o ~/.claude/skills/outlook-web/references/ui-patterns.md \
  https://raw.githubusercontent.com/organisciak/outlook-web-plugin/main/skills/outlook-web/references/ui-patterns.md
curl -o ~/.claude/skills/outlook-web/OPTIMIZATION-AUDIT.md \
  https://raw.githubusercontent.com/organisciak/outlook-web-plugin/main/skills/outlook-web/OPTIMIZATION-AUDIT.md
```

### From Source (Development)

Clone and use locally:

```bash
git clone https://github.com/organisciak/outlook-web-plugin.git
cd outlook-web-plugin
claude --plugin-dir .
```

## Quick Start

### First-Time Setup

Authenticate with Microsoft (one-time only):

```bash
playwright-cli open "https://outlook.office.com/mail/inbox" --session=outlook-web --headed
```

Complete the login in the opened browser window. The session persists for future use.

### Basic Usage

```bash
/outlook-web check my inbox
/outlook-web read the email from John about the meeting
/outlook-web draft a reply saying "Thanks, I'll review this tomorrow"
/outlook-web show my calendar for this week
```

## Usage Examples

### Email

```bash
# Read inbox
/outlook-web show me my unread emails

# Read specific email
/outlook-web open the email from Sarah about project updates

# Compose draft
/outlook-web draft an email to john@example.com with subject "Meeting Follow-up"

# Reply to email
/outlook-web reply to the last email saying "I'll have that ready by Friday"

# Search
/outlook-web search for emails about "budget proposal"
```

### Calendar

```bash
# View calendar
/outlook-web show my calendar for tomorrow

# Check availability
/outlook-web what meetings do I have this afternoon?
```

## Advanced Usage

### Batch Operations

Combine multiple operations for 60-70% faster execution:

```bash
playwright-cli run-code "async () => {
  await page.goto('https://outlook.office.com/mail/inbox');
  await page.getByRole('option').first().click();
  await page.getByRole('button', { name: 'Reply', exact: true }).click();
  await page.getByRole('textbox', { name: 'Message body' }).fill('Your reply');
  await page.keyboard.press('Control+s');
}" --session=outlook-web
```

### Efficient Text Extraction

Extract content without UI overhead (70-90% token reduction):

```bash
# Email list as plain text
playwright-cli run-code "async () => {
  const listbox = await page.locator('[role=\"listbox\"]').first();
  return await listbox.innerText();
}" --session=outlook-web
```

```bash
# Message body only
playwright-cli run-code "async () => {
  const body = await page.locator('[role=\"document\"]').first();
  return await body.innerText();
}" --session=outlook-web
```

### Resource Blocking

Block images, fonts, and analytics for faster page loads:

```bash
playwright-cli run-code "async () => {
  await page.route('**/*', route => {
    const url = route.request().url();
    if (url.match(/\\.(png|jpg|woff|woff2)|browser\\.events/)) {
      route.abort();
    } else {
      route.continue();
    }
  });
  await page.goto('https://outlook.office.com/mail/inbox');
}" --session=outlook-web
```

## Documentation

| File | Description |
|------|-------------|
| [SKILL.md](skills/outlook-web/SKILL.md) | Complete skill reference with all operations |
| [OPTIMIZATION-AUDIT.md](skills/outlook-web/OPTIMIZATION-AUDIT.md) | Performance analysis and optimization strategies |
| [ui-patterns.md](skills/outlook-web/references/ui-patterns.md) | Discovered UI element selectors |

## Safety

| Feature | Description |
|---------|-------------|
| Drafts by default | Emails saved as drafts, never sent without explicit confirmation |
| Session isolation | Dedicated `outlook-web` session avoids conflicts |
| Headless operation | Invisible after initial authentication |
| No credential storage | Uses browser session cookies only |

## Troubleshooting

**Browser already in use:**
```bash
playwright-cli close --session=outlook-web
```

**Session expired (login page appears):**
```bash
playwright-cli open "https://outlook.office.com/mail/inbox" --session=outlook-web --headed
```

**Content not loading:**
```bash
playwright-cli run-code "async () => {
  await page.waitForSelector('[role=\"listbox\"]', { timeout: 10000 });
}" --session=outlook-web
```

## Performance

| Metric | Standard | Optimized | Improvement |
|--------|----------|-----------|-------------|
| Execution time | 11-17s | 2-6s | 60-70% faster |
| Token usage (inbox) | ~7,300 | ~2,000 | 73% reduction |
| Token usage (email) | ~7,500 | ~100 | 99% reduction |

## Compatibility

- **Outlook versions**: outlook.office.com (Office 365/Microsoft 365)
- **Authentication**: Microsoft accounts (work, school, or personal)
- **Browsers**: Chromium-based (via Playwright)
- **Platforms**: macOS, Linux, Windows

## Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Test your changes thoroughly
4. Submit a pull request

## License

MIT License - see [LICENSE](LICENSE) file for details

## Author

Peter Organisciak
- Email: organisciak@gmail.com
- GitHub: [@organisciak](https://github.com/organisciak)

## Acknowledgments

- Built for [Claude Code](https://www.anthropic.com/claude-code) by Anthropic
- Uses [playwright-cli-plugin](https://github.com/anthropics/playwright-cli-plugin)
- Tested with Microsoft Outlook Web (outlook.office.com)

## Support

- **Issues**: [GitHub Issues](https://github.com/organisciak/outlook-web-plugin/issues)
- **Discussions**: [GitHub Discussions](https://github.com/organisciak/outlook-web-plugin/discussions)

---

**Note**: This plugin automates a web interface and may break if Microsoft updates Outlook Web. Please report any issues you encounter.
