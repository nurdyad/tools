from docman.DocmanBaseBot import DocmanBaseBot
from docman.jobs.OnboardingJob import OnboardingJob
from robocorp import workitems


class OnboardingBot(DocmanBaseBot):
    def __init__(self):
        super().__init__()
        self._onboarding_job = OnboardingJob()

    def run_attended(self):
        self._logger.info("Starting onboarding attended mode.")
        self._attended = True

        job = workitems.inputs.current
        practice_id = job.payload["ods_code"]

        # Setup browser and configure job
        self._setup_bot_environment(practice_id)
        self._configure_job(self._onboarding_job)

        # Rebuild Mailroom-like job structure
        mailroom_job = {
            "job": {
                "practice_id": practice_id,
                "parameters": {
                    "user_groups": [
                        "BetterLetter Filing",
                        "BetterLetter Admin",
                        "BetterLetter GPs",
                        "BetterLetter Meds Management",
                        "BetterLetter Safeguarding",
                        "BetterLetter Audit"
                    ],
                    "view_groups": [
                        "BetterLetter Filing",
                        "BetterLetter Rejected",
                        "BetterLetter Processing",
                        "BetterLetter Input"
                    ]
                }
            },
            "attempt_id": job.payload.get("attempt_id", "manual")
        }

        success, error_message, pause_job = self._onboarding_job.process(mailroom_job)
        if not success:
            raise Exception(f"Onboarding failed: {error_message}")

        self._logger.info("Onboarding attended completed successfully.")
