import time
import random
import json
from pathlib import Path
from typing import Dict, Generator, List, Optional
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
    NoSuchElementException,
    StaleElementReferenceException
)

from src.utils.logging import get_logger

logger = get_logger("plm_indexer")

# Default selectors for Aras Innovator - user can override in config
DEFAULT_SELECTORS = {
    # Login page
    "login_form": "form#login, form[name='login'], #loginForm",
    "username_field": "input[name='username'], input[name='user'], input#username, input#user",
    "password_field": "input[name='password'], input[type='password'], input#password",
    "login_button": "button[type='submit'], input[type='submit'], #loginBtn, .login-btn, button:contains('Login')",

    # Session check - if any of these exist, we're logged in
    "logged_in_indicator": ".user-profile, .user-menu, #userMenu, .logout-btn, a[href*='logout'], .aras-nav",

    # Navigation - TOC (Table of Contents) / folder tree
    "toc_container": "#navPanel, .navigation-panel, .toc-tree, #treeView, .folder-tree",
    "toc_folder_items": ".toc-item, .tree-node, .folder-node, li[data-itemtype]",
    "toc_expand_icon": ".expand-icon, .toggle-icon, .tree-toggle, .aras-tree-arrow",

    # Main grid / file list
    "main_grid": "#mainGrid, .search-grid, .item-grid, table.aras-grid, .data-grid",
    "grid_rows": "tr.grid-row, tr[data-id], .grid-item, tbody tr:not(.header)",
    "grid_cell": "td, .grid-cell",

    # Pagination
    "next_page": ".next-page, .pagination-next, a[rel='next'], button:contains('Next')",
    "page_info": ".page-info, .pagination-info, .pager-status",

    # Item details (when clicking into an item)
    "item_name": ".item-name, [data-field='name'], td[data-col='name'], .keyed_name",
    "item_id": ".item-id, [data-field='id'], td[data-col='id'], [data-id]",
    "item_path": ".item-path, [data-field='path'], .breadcrumb",
    "item_created": "[data-field='created_on'], .created-date, td[data-col='created_on']",
    "item_modified": "[data-field='modified_on'], .modified-date, td[data-col='modified_on']",
}


class PLMIndexer:
    def __init__(self, config: Dict):
        self.url = config.get("url")
        self.username = config.get("username")
        self.password = config.get("password")
        self.headless = config.get("headless", True)
        self.save_cookies_flag = config.get("save_cookies", False)
        self.driver = None
        self.cookie_file = Path("config/cookies.json")
        self.wait_timeout = config.get("wait_timeout", 10)
        self.page_load_timeout = config.get("page_load_timeout", 30)

        # Merge user selectors with defaults
        self.selectors = {**DEFAULT_SELECTORS, **config.get("selectors", {})}

        # Folder path to start from (optional)
        self.start_path = config.get("start_path", "/")

        # Track visited folders to avoid loops
        self.visited_folders = set()

        # Track failed paths for partial results
        self.failed_paths: List[str] = []

    def _init_driver(self):
        opts = Options()
        if self.headless:
            opts.add_argument("--headless=new")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1920,1080")
        opts.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})

        # Check for bundled chromedriver
        bundled_driver_paths = [
            Path("bin/chromedriver"),
            Path("bin/chromedriver.exe"),
            Path(__file__).parent.parent.parent / "bin" / "chromedriver",
            Path(__file__).parent.parent.parent / "bin" / "chromedriver.exe"
        ]

        driver_path = None
        for p in bundled_driver_paths:
            if p.exists():
                driver_path = p.resolve()
                break

        service = None
        if driver_path:
            logger.info(f"Using bundled chromedriver at {driver_path}")
            from selenium.webdriver.chrome.service import Service
            service = Service(executable_path=str(driver_path))
        else:
            logger.info("Using system chromedriver (not found in ./bin)")

        try:
            if service:
                self.driver = webdriver.Chrome(service=service, options=opts)
            else:
                self.driver = webdriver.Chrome(options=opts)
            self.driver.set_page_load_timeout(self.page_load_timeout)
        except WebDriverException as e:
            logger.error(f"Failed to initialize Chrome driver: {e}")
            raise

    def _save_cookies(self):
        if not self.driver or not self.save_cookies_flag:
            return
        try:
            cookies = self.driver.get_cookies()
            self.cookie_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.cookie_file, 'w') as f:
                json.dump(cookies, f)
            logger.debug("Session cookies saved.")
        except Exception as e:
            logger.warning(f"Failed to save cookies: {e}")

    def _load_cookies(self) -> bool:
        if not self.driver or not self.cookie_file.exists():
            return False

        try:
            with open(self.cookie_file, 'r') as f:
                cookies = json.load(f)

            if self.url:
                self.driver.get(self.url)

            for cookie in cookies:
                try:
                    self.driver.add_cookie(cookie)
                except Exception:
                    pass
            logger.info("Session cookies loaded.")
            return True
        except Exception as e:
            logger.warning(f"Failed to load cookies: {e}")
            return False

    def _random_sleep(self, min_s: float = 1.0, max_s: float = 3.0):
        time.sleep(random.uniform(min_s, max_s))

    def _find_element(self, selector_key: str, parent=None, timeout: int = None) -> Optional:
        """Find element using comma-separated selector fallbacks."""
        selectors = self.selectors.get(selector_key, "")
        context = parent or self.driver
        timeout = timeout if timeout is not None else self.wait_timeout

        for selector in selectors.split(","):
            selector = selector.strip()
            if not selector:
                continue
            try:
                # Determine selector type
                if selector.startswith("//"):
                    by = By.XPATH
                elif selector.startswith("#"):
                    by = By.CSS_SELECTOR
                elif selector.startswith("."):
                    by = By.CSS_SELECTOR
                else:
                    by = By.CSS_SELECTOR

                if timeout > 0:
                    element = WebDriverWait(context, timeout).until(
                        EC.presence_of_element_located((by, selector))
                    )
                else:
                    element = context.find_element(by, selector)
                return element
            except (TimeoutException, NoSuchElementException):
                continue
        return None

    def _find_elements(self, selector_key: str, parent=None) -> List:
        """Find all elements using comma-separated selector fallbacks."""
        selectors = self.selectors.get(selector_key, "")
        context = parent or self.driver

        for selector in selectors.split(","):
            selector = selector.strip()
            if not selector:
                continue
            try:
                if selector.startswith("//"):
                    by = By.XPATH
                else:
                    by = By.CSS_SELECTOR

                elements = context.find_elements(by, selector)
                if elements:
                    return elements
            except Exception:
                continue
        return []

    def _is_logged_in(self) -> bool:
        """Check if we're already logged in."""
        indicator = self._find_element("logged_in_indicator", timeout=3)
        return indicator is not None

    def login(self):
        if not self.driver:
            self._init_driver()

        # Try cookie reuse first
        if self._load_cookies():
            self.driver.get(self.url)
            self._random_sleep(2, 4)
            if self._is_logged_in():
                logger.info("Restored session from cookies.")
                return

        logger.info(f"Navigating to {self.url}...")
        self.driver.get(self.url)
        self._random_sleep(2, 4)

        # Check if already logged in (SSO, cached session)
        if self._is_logged_in():
            logger.info("Already logged in.")
            self._save_cookies()
            return

        # Find and fill login form
        login_form = self._find_element("login_form", timeout=10)
        if not login_form:
            logger.warning("Login form not found. May need manual login or different selectors.")
            logger.info("Waiting 30s for manual login if running non-headless...")
            time.sleep(30)
            self._save_cookies()
            return

        # Fill credentials
        username_field = self._find_element("username_field")
        password_field = self._find_element("password_field")

        if username_field and self.username:
            username_field.clear()
            username_field.send_keys(self.username)
            self._random_sleep(0.5, 1)

        if password_field and self.password:
            password_field.clear()
            password_field.send_keys(self.password)
            self._random_sleep(0.5, 1)

        # Submit
        login_button = self._find_element("login_button")
        if login_button:
            login_button.click()
            logger.info("Login form submitted.")
        else:
            # Try form submit
            if password_field:
                from selenium.webdriver.common.keys import Keys
                password_field.send_keys(Keys.RETURN)
                logger.info("Login form submitted via Enter key.")

        # Wait for login to complete
        self._random_sleep(3, 5)

        # Check for MFA prompt or additional steps
        # This is highly site-specific - log current state for debugging
        if not self._is_logged_in():
            logger.warning("Login may require MFA or additional steps.")
            logger.info(f"Current URL: {self.driver.current_url}")
            logger.info("Waiting 60s for manual intervention...")
            time.sleep(60)

        self._save_cookies()
        logger.info("Login complete.")

    def _extract_item_data(self, row_element, current_path: str) -> Optional[Dict]:
        """Extract item metadata from a grid row."""
        try:
            # Try to get item name
            name = None
            name_elem = self._find_element("item_name", parent=row_element, timeout=1)
            if name_elem:
                name = name_elem.text.strip()

            # Fallback: first cell often contains name
            if not name:
                cells = self._find_elements("grid_cell", parent=row_element)
                if cells:
                    name = cells[0].text.strip()

            if not name:
                return None

            # Get ID
            item_id = None
            id_elem = self._find_element("item_id", parent=row_element, timeout=1)
            if id_elem:
                item_id = id_elem.text.strip() or id_elem.get_attribute("data-id")

            # Fallback: check row's data-id attribute
            if not item_id:
                item_id = row_element.get_attribute("data-id") or row_element.get_attribute("id")

            # Get dates
            created_at = None
            modified_at = None

            created_elem = self._find_element("item_created", parent=row_element, timeout=1)
            if created_elem:
                created_at = created_elem.text.strip()

            modified_elem = self._find_element("item_modified", parent=row_element, timeout=1)
            if modified_elem:
                modified_at = modified_elem.text.strip()

            # Build remote path
            remote_path = f"{current_path}/{name}".replace("//", "/")

            return {
                "name": name,
                "remote_path": remote_path,
                "remote_id": item_id,
                "created_at": created_at or datetime.now().isoformat(),
                "modified_at": modified_at or datetime.now().isoformat(),
                "source": "plm"
            }
        except StaleElementReferenceException:
            logger.debug("Stale element encountered, skipping row")
            return None
        except Exception as e:
            logger.debug(f"Failed to extract item data: {e}")
            return None

    def _get_grid_items(self, current_path: str) -> Generator[Dict, None, None]:
        """Extract all items from the current grid view."""
        page_num = 1

        while True:
            logger.debug(f"Processing page {page_num} of {current_path}")
            self._random_sleep(1, 2)

            # Wait for grid to load
            grid = self._find_element("main_grid", timeout=10)
            if not grid:
                logger.warning(f"Grid not found at {current_path}")
                break

            # Get all rows
            rows = self._find_elements("grid_rows")
            if not rows:
                logger.debug(f"No rows found at {current_path}")
                break

            logger.info(f"Found {len(rows)} rows on page {page_num}")

            for row in rows:
                item_data = self._extract_item_data(row, current_path)
                if item_data:
                    yield item_data
                self._random_sleep(0.1, 0.3)  # Small delay between rows

            # Check for pagination
            next_button = self._find_element("next_page", timeout=2)
            if next_button:
                try:
                    # Check if button is disabled
                    if next_button.get_attribute("disabled") or "disabled" in (next_button.get_attribute("class") or ""):
                        break
                    next_button.click()
                    page_num += 1
                    self._random_sleep(2, 4)  # Longer pause for page loads
                except Exception as e:
                    logger.debug(f"Pagination ended: {e}")
                    break
            else:
                break

            # Safety limit
            if page_num > 100:
                logger.warning("Reached page limit (100), stopping pagination")
                break

    def _navigate_to_folder(self, folder_path: str) -> bool:
        """Navigate to a specific folder in the TOC/tree view."""
        try:
            # If we have a direct URL pattern, use it
            if "/folder/" in self.url or "ItemType=Folder" in self.url:
                # Construct folder URL (site-specific)
                folder_url = f"{self.url}?path={folder_path}"
                self.driver.get(folder_url)
                self._random_sleep(2, 4)
                return True

            # Otherwise, try to navigate via TOC
            toc = self._find_element("toc_container", timeout=5)
            if not toc:
                logger.warning("TOC/Navigation panel not found")
                return False

            # Split path and navigate
            parts = [p for p in folder_path.split("/") if p]
            for part in parts:
                # Find folder in tree
                folder_items = self._find_elements("toc_folder_items")
                found = False
                for item in folder_items:
                    if part.lower() in item.text.lower():
                        # Expand if needed
                        expand_icon = self._find_element("toc_expand_icon", parent=item, timeout=1)
                        if expand_icon:
                            expand_icon.click()
                            self._random_sleep(0.5, 1)
                        item.click()
                        self._random_sleep(1, 2)
                        found = True
                        break

                if not found:
                    logger.warning(f"Folder '{part}' not found in navigation")
                    return False

            return True
        except Exception as e:
            logger.error(f"Failed to navigate to {folder_path}: {e}")
            return False

    def scan(self) -> Generator[Dict, None, None]:
        """Scrape the PLM interface for files."""
        if not self.driver:
            self.login()

        logger.info("Starting PLM scan...")
        self._random_sleep(2, 5)

        # Start from configured path or root
        folders_to_process = [self.start_path]
        items_yielded = 0

        while folders_to_process:
            current_folder = folders_to_process.pop(0)

            if current_folder in self.visited_folders:
                continue
            self.visited_folders.add(current_folder)

            logger.info(f"Scanning folder: {current_folder}")

            try:
                # Navigate to folder
                if not self._navigate_to_folder(current_folder):
                    self.failed_paths.append(current_folder)
                    continue

                # Process items in this folder
                for item in self._get_grid_items(current_folder):
                    yield item
                    items_yielded += 1

                    # Longer pause every N items
                    if items_yielded % 50 == 0:
                        logger.info(f"Processed {items_yielded} items...")
                        self._random_sleep(3, 6)

                # Find subfolders to process (if any)
                # This is highly site-specific - some PLMs show folders in the grid
                # others have a separate tree view
                subfolder_items = self._find_elements("toc_folder_items")
                for item in subfolder_items:
                    try:
                        folder_name = item.text.strip()
                        if folder_name:
                            subfolder_path = f"{current_folder}/{folder_name}".replace("//", "/")
                            if subfolder_path not in self.visited_folders:
                                folders_to_process.append(subfolder_path)
                    except Exception:
                        pass

            except Exception as e:
                logger.error(f"Error processing folder {current_folder}: {e}")
                self.failed_paths.append(current_folder)
                continue

        logger.info(f"PLM scan complete. Total items: {items_yielded}")
        if self.failed_paths:
            logger.warning(f"Failed to process {len(self.failed_paths)} folders: {self.failed_paths}")

    def close(self):
        if self.driver:
            self._save_cookies()
            self.driver.quit()
            self.driver = None
