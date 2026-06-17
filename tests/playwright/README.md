# Playwright E2E

```bash
npm init -y
npm i -D @playwright/test
npx playwright install chromium

# Start the app in another terminal:
python3 app.py

# Then:
npx playwright test --config tests/playwright/playwright.config.js
```

Set `BANKI_URL` if the app runs on a non-default host/port.
