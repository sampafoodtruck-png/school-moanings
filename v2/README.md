# School Moanings — Version 2 draft

This is an additive preview under `/v2/`. The original `index.html` is unchanged.

## Behaviour

- Familiar morning and evening routines, plus Today, Prepare for tomorrow, and a seven-day week view.
- Read-only personal Outlook calendar access. Automatically refreshes every 15 minutes while this page is open and the Microsoft session remains valid. The page shows the last refresh and warns about stale or uncovered dates. It cannot refresh while the browser is closed and Microsoft may require another sign-in.
- Reads concrete recurring occurrences and exceptions via Microsoft Graph calendarView, following every result page. Cancelled and private/confidential entries are excluded. The view defaults to Brisbane time, regardless of the display device's timezone.
- Only subjects matching the activity words are retained in the local calendar copy. Routine tasks and activity packing ticks are separate; packing ticks do not alter the original pizza progress.
- Swimming and gymnastics packing suggestions are explicitly marked as suggestions. Surf Skills instructions are drawn from the event notes when those notes say to wear swimming gear. Entries without a child's name say “Check who is attending”; they are not silently assigned to both children.
- Main calendar is the default. After sign-in, choose a different available calendar in Calendar settings if needed. Save and refresh after switching calendars or activity words.
- No calendar events, account identifiers, tokens or secrets are committed to the repository. Events and packing ticks are saved only in this browser's local storage; authentication uses MSAL session storage. Every device signs in separately; cross-device tick sync is not implemented.

## One-time Microsoft setup

Connecting Outlook inside ChatGPT does not give this static website an OAuth application ID or Microsoft token. A separate Microsoft app registration is required.

1. Register an application in Microsoft Entra: https://entra.microsoft.com/ (App registrations → New registration).
2. Choose support for **personal Microsoft accounts** (or organisational and personal accounts). Use delegated access, not application permissions.
3. Under Authentication, add a **Single-page application** redirect URI:
   `https://sampafoodtruck-png.github.io/school-moanings/v2/auth.html`
   This exact address applies after the draft is merged and GitHub Pages publishes it. For another preview host, register that host's exact `/v2/auth.html` address instead.
4. Add Microsoft Graph **delegated `Calendars.Read`** permission. This permission is needed for the event notes used by the packing rules. Do not add calendar write permissions. Do not create a client secret or enable implicit token grants.
5. Copy the **Application (client) ID** into Calendar settings on the dashboard. It is an identifier, not a password. Sign in to Outlook and approve read access.
6. Choose the intended calendar, adjust activity words and refresh. Check a known swimming event and its date before relying on the board.

If app registration is unavailable for the personal account, stop here and choose an authenticated server integration with the account owner. Do not publish a private calendar feed or tokens as a workaround.

The Microsoft authentication module is loaded on demand from the pinned npm ESM package `@azure/msal-browser@4.30.0` via jsDelivr. Microsoft lists the v4 line as LTS in the package release tags. No credentials are sent to jsDelivr by this code. A release can vendor/bundle the official npm package to remove that external script dependency.

## Verification

Run `node v2/tests/calendar.test.cjs` from the repository root. The dependency-free tests cover Brisbane boundaries, exclusive all-day end dates, cancelled/private filtering, explicit child matching, HTML escaping, packing rules, refresh pagination and stale-data preservation after a failed refresh. JavaScript is also parsed for syntax.

Microsoft sign-in, consent, popup callback, Graph permissions and visual layout still require a real browser test with a registered application. These are not claimed as end-to-end verified. MSAL and Graph cannot be exercised with this chat's connector credentials.

References:
- https://learn.microsoft.com/en-us/entra/msal/javascript/browser/initialization
- https://learn.microsoft.com/en-us/graph/api/calendar-list-calendarview?view=graph-rest-1.0
- https://learn.microsoft.com/en-us/graph/permissions-reference#calendarsread
