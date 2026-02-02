# Outlook Web Skill - Optimization Audit

## Reply Flow Audit

### Current Flow (VERIFIED WORKING ✅)

**Steps to draft a reply:**
1. Navigate to inbox: `playwright-cli open "https://outlook.office.com/mail/inbox"`
2. Click email: `playwright-cli click <email-ref>`
3. Click Reply button: `playwright-cli click <reply-button-ref>`
4. Fill message body: `playwright-cli fill <body-ref> "Reply text"`
5. Save draft: `playwright-cli press "Control+s"` or `playwright-cli press Escape`
6. Verify in drafts: `playwright-cli open "https://outlook.office.com/mail/drafts"`

**Total operations:** 6 commands
**Success:** Draft created with "Re: [Subject]" format, auto-saved ✅

### Observations

**What works well:**
- Session persistence keeps authentication between commands
- Direct URL navigation is fast and reliable
- Auto-save happens automatically (no explicit save needed in most cases)
- Headless mode reduces overhead vs headed browser
- Email URLs include unique IDs for direct access

**Current bottlenecks:**
- Multiple page loads (inbox → email → drafts for verification)
- Full page rendering even though we only need specific elements
- Snapshots include entire DOM tree (large YAML files)
- Network loads 100+ resources per page (JS, images, analytics, fonts, etc.)
- No resource blocking configured (images, fonts, analytics all load)

## Programmatic Optimizations

### 1. Resource Blocking (HIGH IMPACT 🚀)

Playwright supports blocking resource types to speed up page loads. However, `playwright-cli` doesn't expose this directly via command line.

**What could be blocked:**
- Images (`.png`, `.jpg`, `.svg`)
- Fonts (`.woff`, `.woff2`, `.ttf`)
- Analytics/telemetry (`browser.events.data.microsoft.com`, `csp.microsoft.com`)
- Stylesheets (some CSS if not critical)
- Media files

**Potential implementation:**
- Create a `playwright-cli.json` config file with custom Playwright context options
- Use `run-code` command to execute custom Playwright code for route blocking
- Estimated speed improvement: **30-50% faster page loads**

**Example route blocking code:**
```javascript
await page.route('**/*', (route) => {
  const url = route.request().url();
  if (url.includes('browser.events.data.microsoft.com') ||
      url.includes('csp.microsoft.com') ||
      url.endsWith('.png') || url.endsWith('.jpg') ||
      url.endsWith('.woff') || url.endsWith('.woff2')) {
    route.abort();
  } else {
    route.continue();
  }
});
```

### 2. Direct Email ID Navigation (MEDIUM IMPACT 📈)

**Current:** Navigate to inbox, find email, click
**Optimized:** Navigate directly to email by ID

Email URLs have format: `https://outlook.office.com/mail/inbox/id/<EMAIL_ID>`

**Benefits:**
- Skip inbox page load
- Skip email list rendering
- Skip email search/click interaction

**Challenge:**
- Need to extract/store email IDs from initial reads
- Email IDs are long encoded strings (Base64-like)

**Recommendation:** For workflows where you know the email ID (e.g., from a previous read or notification), navigate directly.

### 3. Minimize Snapshots (LOW-MEDIUM IMPACT 📊)

**Current:** Snapshots after every operation for verification
**Optimized:** Snapshots only when needed

**Strategy:**
- Skip snapshots for intermediate steps
- Only snapshot for final verification or when extracting data
- Trust command success unless verification is critical

**Example optimized flow:**
```bash
# No snapshot after each step
playwright-cli open "https://outlook.office.com/mail/inbox" --session=outlook-web
playwright-cli click e600 --session=outlook-web
playwright-cli click e1219 --session=outlook-web  # Reply button
playwright-cli fill e2181 "Reply text" --session=outlook-web
playwright-cli press "Control+s" --session=outlook-web
# Only snapshot at end if needed
playwright-cli snapshot --session=outlook-web
```

### 4. Batch Operations with run-code (MEDIUM IMPACT 🔧)

Instead of multiple CLI commands, use `run-code` to batch operations:

```bash
playwright-cli run-code "
  await page.goto('https://outlook.office.com/mail/inbox');
  await page.getByRole('option', { name: 'Unread Collapsed Shimelis' }).click();
  await page.getByRole('button', { name: 'Reply' }).click();
  await page.getByRole('textbox', { name: 'Message body' }).fill('Reply text');
  await page.keyboard.press('Control+s');
" --session=outlook-web
```

**Benefits:**
- Single process invocation
- Reduced IPC overhead between commands
- Faster execution
- Can add custom logic (conditionals, loops)

### 5. Use Faster Selectors (LOW IMPACT ⚡)

**Current:** Using ref IDs from snapshots (e.g., `e600`)
**Alternative:** Direct role-based selectors

Ref IDs require snapshot parsing, while direct selectors can be used immediately:

```bash
# Slower (requires snapshot first)
playwright-cli snapshot
playwright-cli click e600

# Faster (direct selector)
playwright-cli click "[aria-label='Reply']"
```

**Tradeoff:** Direct selectors may be less reliable if UI changes, but faster.

### 6. Session Configuration Optimization (LOW IMPACT ⚙️)

Create a `playwright-cli.json` config for the outlook-web session:

```json
{
  "contextOptions": {
    "viewport": { "width": 1280, "height": 720 },
    "userAgent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
    "locale": "en-US",
    "timezoneId": "America/Denver",
    "reducedMotion": "reduce",
    "serviceWorkers": "block"
  },
  "launchOptions": {
    "args": [
      "--disable-blink-features=AutomationControlled",
      "--disable-dev-shm-usage",
      "--no-sandbox"
    ]
  }
}
```

**Benefits:**
- `reducedMotion: reduce` - Less animation overhead
- `serviceWorkers: block` - Skip service worker registration
- Smaller viewport = less rendering

### 7. Skip Unnecessary Navigations (HIGH IMPACT 🎯)

**Current:** Navigate to drafts to verify
**Optimized:** Trust auto-save, skip verification

Outlook Web auto-saves drafts every few seconds. In programmatic use:
- Skip draft verification unless critical
- Assume draft saved if no error occurred
- Only verify when user explicitly requests

## Snapshot Efficiency for LLM Consumption

### Current Snapshot Structure

Playwright-cli generates snapshots using the **accessibility tree** (not raw DOM), which is already a significant optimization:

**What IS included (Good):**
- Semantic roles (button, textbox, option, listbox)
- ARIA labels and accessible names
- Text content
- Interactive element references
- Hierarchical structure

**What is NOT included (Good):**
- Raw HTML tags (div, span, etc.)
- CSS classes and IDs
- Inline styles
- JavaScript code
- Hidden/aria-hidden elements
- Script and style tags

### The Problem: Signal-to-Noise Ratio

**Inbox snapshot analysis (731 lines):**
```
Lines 1-90:   Banner, app launcher, search (12%)
Lines 91-362: Ribbon toolbar (Home tab buttons) (37%)
Lines 363-437: Navigation pane, folder tree (10%)
Lines 438-460: Message list header (3%)
Lines 461-730: Actual email list (13 emails) (37%)

Element breakdown:
- 324 "generic" wrappers (44%)
- 118 interactive elements (16%)
- 289 other semantic elements (40%)

Content vs Chrome:
- Email content: 270 lines (37%)
- UI chrome: 460 lines (63%)
```

**Reading pane snapshot (749 lines):**
```
Lines 1-720:  UI chrome (toolbar, nav, email metadata) (96%)
Lines 721-726: Actual message body content (5 lines) (0.7%)

For a 5-paragraph email, only 5 lines contain the actual message text!
```

### Why This Happens

Outlook Web is a **single-page application** with complex UI:
- Persistent navigation (folder tree, app launcher)
- Feature-rich ribbon with 50+ buttons
- Email metadata (sender, recipient, timestamp, actions)
- Nested generic wrappers for styling/layout

The accessibility tree correctly represents **everything visible**, but for programmatic/LLM use, most of this is noise.

### Impact on LLM Processing

**Token usage:**
- Full inbox snapshot: ~731 lines × ~10 tokens/line = **~7,300 tokens**
- Actual email list: ~270 lines = **~2,700 tokens**
- Wasted: **~4,600 tokens per snapshot** on UI chrome

**Latency:**
- LLM must parse 460 lines before finding first email
- Must scan through generic wrappers to find semantic content
- Pattern matching becomes harder with noise

### Optimization Strategies

#### 1. Extract Text with JavaScript Functions (BEST OPTION 🌟)

Use playwright-cli's code execution to extract only content you need:

```bash
# Get just the email list text (without toolbar/nav)
playwright-cli run-code "
  const listbox = await page.locator('[role=\"listbox\"]');
  return await listbox.innerText();
" --session=outlook-web
```

```bash
# Get just the message body text
playwright-cli run-code "
  const body = await page.locator('[role=\"document\"]').first();
  return await body.innerText();
" --session=outlook-web
```

**Benefits:**
- Extract only text content (no YAML structure overhead)
- Skip all UI chrome
- Reduces token usage by 70-90%

#### 2. Parse Snapshots Intelligently

When consuming snapshots, skip irrelevant sections:

```python
# Example: Extract only email list from snapshot
lines = snapshot.split('\n')
email_list_start = next(i for i, line in enumerate(lines) if 'Message list' in line)
email_content = lines[email_list_start:]
```

**Benefits:**
- Reuse existing snapshots
- Filter programmatically
- Reduces tokens sent to LLM

#### 3. Direct API Access (ADVANCED)

Instead of UI scraping, use Outlook REST API:

```bash
# Requires authentication, but much cleaner
curl -H "Authorization: Bearer $TOKEN" \
  "https://outlook.office.com/api/v2.0/me/messages?$top=10&$select=subject,from,bodyPreview"
```

**Benefits:**
- Clean JSON response
- No UI chrome at all
- Much faster
- Standard API

**Drawbacks:**
- Requires OAuth setup
- Separate auth from web session
- More complex for simple use

#### 4. Targeted Extraction with run-code

Combine navigation + extraction in one command:

```bash
playwright-cli run-code "
  await page.goto('https://outlook.office.com/mail/inbox');
  await page.waitForSelector('[role=\"option\"]');

  // Extract just what you need
  const emails = await page.locator('[role=\"option\"]')
    .evaluateAll(options =>
      options.slice(0, 10).map(opt => ({
        text: opt.textContent.trim(),
        unread: opt.textContent.includes('Unread')
      }))
    );

  return JSON.stringify(emails, null, 2);
" --session=outlook-web
```

**Benefits:**
- Custom data extraction
- JSON output (easier to parse)
- No snapshot overhead

### Recommendations by Use Case

**Casual/Interactive Use:**
- Use full snapshots (current approach)
- Human-readable YAML format
- Visual inspection of UI state
- **Token cost:** ~7k per snapshot

**Programmatic Email Reading:**
- Use run-code to extract message body text
- Skip snapshots for navigation
- **Token cost:** ~100-500 per extraction

**Bulk Email Processing:**
- Use run-code for batch extraction
- Return JSON arrays
- Consider Outlook REST API
- **Token cost:** ~50-100 per email

**Hybrid (Recommended for LLMs):**
- Snapshot once to understand structure
- Use run-code with selectors to extract content
- Cache selectors for repeated access
- **Token cost:** ~1k initial + ~100 per extraction

## Performance Comparison

### Current Flow (6 commands)
```
open inbox       →  ~3-5s (full page load)
snapshot         →  ~0.5s
click email      →  ~2-3s (content load)
snapshot         →  ~0.5s
click reply      →  ~1-2s (compose opens)
fill body        →  ~0.3s
save draft       →  ~0.2s
open drafts      →  ~3-5s (verification)
snapshot         →  ~0.5s
─────────────────────────────
TOTAL:           ~11-17s
```

### Optimized Flow (3-4 commands, with blocking)
```
run-code batch   →  ~2-4s (with resource blocking)
  ↳ open inbox + click email + reply + fill + save
[optional verify] →  ~1-2s
─────────────────────────────
TOTAL:            ~2-6s (60-70% faster)
```

## Recommended Optimizations by Priority

### 🥇 HIGH PRIORITY (Implement First)
1. **Resource blocking** - Biggest impact on load times
2. **Skip draft verification** - Trust auto-save in automated workflows
3. **Batch with run-code** - Reduce command overhead

### 🥈 MEDIUM PRIORITY (Quick Wins)
4. **Direct email ID navigation** - When ID is known
5. **Minimize snapshots** - Only when extracting data

### 🥉 LOW PRIORITY (Diminishing Returns)
6. **Session config** - Minor improvements
7. **Faster selectors** - Marginal gains, less reliable

## Implementation Notes

**Resource blocking requires custom code:**
- Not available via command-line flags
- Must use `run-code` or custom config
- May need to maintain allow/block lists

**Trade-offs:**
- Speed vs Reliability: Fewer verifications = faster but less safe
- Simplicity vs Performance: Command-line is simple; run-code is faster but more complex
- Maintenance: Resource blocking needs updates if Outlook Web changes CDNs/domains

## Conclusion

The current skill is **functional and reasonably efficient** for general use. For **high-volume programmatic use**, implementing resource blocking and batching via `run-code` could reduce execution time by **60-70%**.

The biggest optimization opportunity is **resource blocking**, but it requires writing custom Playwright code, which adds complexity to the skill.

## Tested Results

### Email List Extraction ✅

**Full snapshot:**
```
731 lines of YAML
~7,300 tokens
63% UI chrome, 37% email content
```

**Text extraction with run-code:**
```
~200 lines of plain text
~2,000 tokens
100% email content
73% token reduction
```

### Email Body Extraction ✅

**Full snapshot:**
```
749 lines of YAML
~7,500 tokens
96% UI chrome, 0.7% message content
```

**Text extraction with run-code:**
```
~10 lines of plain text
~100 tokens
100% message content
99% token reduction
```

**Recommendation:**
- Keep current approach for general use (simple, reliable)
- Add text extraction examples for LLM/programmatic use
- Document resource blocking as an optional optimization for power users
