# skill/switch_merchant/logic.py

import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from common.logger import get_logger

log = get_logger("merchant_switcher")

def auto_switch_merchant(driver, target_name):
    """
    Automated merchant switch using the profile menu dropdown.
    
    Args:
        driver: Selenium WebDriver instance
        target_name (str): The name of the merchant to switch to (Case Sensitive, as shown in Shopee menu)
    
    Returns:
        bool: True if switch was successful, False otherwise.
    """
    log.info(f"🔄 Attempting auto-switch to: {target_name}...")
    try:
        # Ensure we are on a page where the menu is available
        if "/food/dashboard" not in driver.current_url:
            driver.get("https://partner.shopee.co.id/food/dashboard")
            time.sleep(2)
        
        wait = WebDriverWait(driver, 15)
        actions = ActionChains(driver)
        
        # 1. Locate Profile/Account menu
        selectors = [
            "li[data-menu-id*='account']", 
            ".ant-menu-item-only-child[data-menu-id*='account']", 
            "li.ant-menu-item:last-child"
        ]
        profile_menu = None
        for sel in selectors:
            try:
                profile_menu = driver.find_element(By.CSS_SELECTOR, sel)
                break
            except: 
                continue
            
        if not profile_menu:
            raise Exception("Could not find Profile/Account menu")

        # Scroll into view and hover
        driver.execute_script("arguments[0].scrollIntoView(true);", profile_menu)
        time.sleep(0.3)
        actions.move_to_element(profile_menu).perform()
        time.sleep(0.5) 
        
        # 2. Locate and hover "Pilih Merchant Lain"
        try:
            switch_trigger = wait.until(EC.presence_of_element_located((By.XPATH, "//span[text()='Pilih Merchant Lain']")))
            actions.move_to_element(switch_trigger).perform()
            time.sleep(0.5)
        except Exception as e:
            log.warning(f"Could not hover 'Pilih Merchant Lain': {e}")
            # Try finding it in the DOM directly if hover fails
            pass
        
        # 3. Use JavaScript to find and click the target merchant name
        # This is more robust against scrolling and dynamic menus
        log.info(f"  Searching for '{target_name}' using JS scanner...")
        js_click_script = f"""
            var targetName = "{target_name.lower()}";
            var spans = document.querySelectorAll('span, div, li');
            for (var i = 0; i < spans.length; i++) {{
                var text = (spans[i].innerText || "").toLowerCase().trim();
                if (text === targetName || (text.includes(targetName) && spans[i].children.length === 0)) {{
                    spans[i].scrollIntoView({{block: 'center'}});
                    spans[i].click();
                    return true;
                }}
            }}
            return false;
        """
        
        success = False
        for _ in range(5):
            if driver.execute_script(js_click_script):
                success = True
                break
            # If not found, try scrolling the menu containers if they exist
            driver.execute_script("document.querySelectorAll('div[class*=\"menu\"], ul[class*=\"menu\"]').forEach(el => el.scrollTop += 500);")
            time.sleep(1)
            
        if not success:
            raise Exception(f"JS Scanner could not find merchant: {target_name}")
            
        log.info(f"✅ Clicked {target_name}. Waiting for dashboard redirect...")
        wait.until(EC.url_contains("/food/dashboard"))
        time.sleep(1) 
        return True
    except Exception as e:
        log.error(f"❌ Auto-switch failed: {e}")
        return False

def return_to_selector(driver):
    """
    Navigates back to the merchant selector page to allow manual switching.
    
    Args:
        driver: Selenium WebDriver instance
        
    Returns:
        bool: True if successful, False otherwise.
    """
    log.info("🔄 Returning to merchant selector page...")
    selector_url = "https://partner.shopee.co.id/authenticate/merchant-selector"
    driver.get(selector_url)
    
    try:
        wait = WebDriverWait(driver, 30)
        wait.until(lambda d: "/merchant-selector" in d.current_url or "/food/dashboard" in d.current_url)
        return True
    except Exception as e:
        log.error(f"❌ Failed to reach merchant selector page: {e}")
        return False

def get_current_merchant_name_from_ui(driver):
    """
    Reads the active merchant name from the dashboard UI.
    
    Args:
        driver: Selenium WebDriver instance
        
    Returns:
        str or None: The merchant name if found, else None.
    """
    try:
        wait = WebDriverWait(driver, 10)
        # The class 'merchantName' is commonly used in Shopee Partner dashboard
        name_el = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "merchantName")))
        return name_el.text.strip()
    except:
        # Fallback to other possible selectors
        try:
            name_el = driver.find_element(By.CSS_SELECTOR, ".ant-dropdown-trigger .name")
            return name_el.text.strip()
        except:
            return None
