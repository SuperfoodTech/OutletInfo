import asyncio
import logging
from playwright.async_api import async_playwright

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("TestGrab")

TEST_USERNAME = "automationf1"
TEST_PASSWORD = "Automationf1@"
LOGIN_URL = "https://merchant.grab.com/login"
DASHBOARD_URL = "https://merchant.grab.com/dashboard"


async def test_login():
    """Test apakah login ke Grab Merchant berhasil."""
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        logger.info(f"Navigating to {LOGIN_URL}")
        await page.goto(LOGIN_URL, wait_until="networkidle", timeout=60000)

        try:
            username_input = page.locator('input[type="text"]').first
            await username_input.fill(TEST_USERNAME)
            logger.info("Username diisi")

            password_input = page.locator('input[type="password"]').first
            await password_input.fill(TEST_PASSWORD)
            logger.info("Password diisi")

            submit_btn = page.locator('button[type="submit"]').first
            await submit_btn.click()
            logger.info("Submit diklik")

            await page.wait_for_url(f"https://merchant.grab.com/**", timeout=30000)
            logger.info(f"[✓] Login berhasil! URL sekarang: {page.url}")

        except Exception as e:
            logger.error(f"[!] Login gagal: {e}")

        await page.wait_for_timeout(3000)
        await browser.close()


if __name__ == "__main__":
    asyncio.run(test_login())
