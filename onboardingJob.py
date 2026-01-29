from docman.DocmanBaseJob import DocmanBaseJob
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from time import sleep
import traceback


class OnboardingJob(DocmanBaseJob):
    def __init__(self):
        super().__init__()

    def _job_specific_process(self, job):
        try:
            ods_code = job["job"]["practice_id"]
            self._logger.info(f"Starting Docman onboarding for ODS code: {ods_code}")

            self._create_folders()
            self._create_user_groups(job["job"]["parameters"]["user_groups"])
            self._create_views(job["job"]["parameters"]["view_groups"])
            self._configure_search_settings()
            
            self._logger.info(f"Docman onboarding complete for ODS code: {ods_code}")
            return True, None, False
        
        except Exception as e:
            self._logger.error(f"Error during Docman onboarding: {e}")
            self._logger.error(traceback.format_exc())
            return False, str(e), False

    def _create_folders(self):
        self._logger.info("Creating folders...")
        folders = [
            "BetterLetter: Filing",
            "BetterLetter: Rejected",
            "BetterLetter: Processing",
            "BetterLetter: Input",
        ]

        self._browser.click("a:has-text('Settings')")
        self._browser.click('span:has-text("Filing")')
        self._browser.click('a:has-text("Document Folders")')
        self._browser.click('td >> a:has-text("Filing")')
        self._browser.click('a:has-text("Top Level Folder")')

        for folder in folders:
            try:
                self._browser.click("a#addFolder")
                textbox = self._browser.locator("input#txtNewFolderName")
                textbox.fill("")
                sleep(0.2)
                for char in folder:
                    textbox.type(char)
                    sleep(0.05)

                self._browser.wait_for_selector("a#addFolderConfirm:not([disabled])", timeout=8000)
                self._browser.click("a#addFolderConfirm")

                # Handle duplicate modal
                try:
                    modal = self._browser.wait_for_selector("text=A folder with the name", timeout=3000)
                    if modal:
                        self._logger.warning(f"Duplicate detected: {folder} — clicking 'No'")
                        self._browser.click("text=No")
                        continue
                except PlaywrightTimeoutError:
                    pass

                self._logger.info(f"Created folder: {folder}")
            except Exception as e:
                self._logger.warning(f"Could not create folder '{folder}': {e}")

        # ✅ Fix for navigation hang
        try:
            self._logger.info("Clicking 'Back to application'")
            self._browser.click("a:has-text('Back to application')")
        except PlaywrightTimeoutError:
            self._logger.warning("Could not find 'Back to application' link. Proceeding anyway.")



    def _create_user_groups(self, groups_to_create):
        self._logger.info("Creating user groups...")
        
        self._browser.click(selector="a:has-text('Settings')")
        self._browser.click(selector='span:has-text("Users")')
        self._browser.click(selector='a:has-text("User Groups")')
        
        for group_name in groups_to_create:
            self._browser.click(selector="a:has-text('Create')")
            self._browser.fill(selector="input#group_name_input", value=group_name)
            self._browser.click(selector="a:has-text('Confirm')")
            self._logger.info(f"Created user group: {group_name}")

    def _create_views(self, views_to_create):
        self._logger.info("Creating views...")
        
        self._browser.click(selector="a:has-text('Settings')")
        self._browser.click(selector='span:has-text("Tasks")')
        self._browser.click(selector='a:has-text("Views")')
        
        for view_name in views_to_create:
            self._browser.click(selector="a:has-text('Create New View')")
            self._browser.fill(selector="input#view_name_input", value=view_name)
            self._browser.fill(selector="input#available_to_input", value="Everyone")
            
            self._browser.click(selector='//select[@id="sent_to_select"]')
            self._browser.press(selector='//select[@id="sent_to_select"]', key="ArrowDown")
            self._browser.press(selector='//select[@id="sent_to_select"]', key="Enter")

            self._browser.click(selector='//input[@id="sent_to_group_select"]')
            self._browser.fill(selector='//input[@id="sent_to_group_select"]', value=view_name)
            self._browser.press(selector='//input[@id="sent_to_group_select"]', key="Enter")
            
            self._browser.click(selector="button#confirm_create_view")
            self._logger.info(f"Created view: {view_name}")

    def _configure_search_settings(self):
        self._logger.info("Configuring search settings...")
        self._browser.click(selector="a:has-text('Settings')")
        self._browser.click(selector='a:has-text("My profile")')
        self._browser.click(selector='a:has-text("Search settings")')
        
        self._select_in_select2("div#s2id_dm-search-in", "select2-results", "patient")
        self._select_in_select2("div#s2id_dm-search-using", "select2-results", "Name, DOB or NHS")
        
        self._browser.click(selector='label[for="hide_synthetic_patients"]')