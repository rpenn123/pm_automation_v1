# Release checklist

# Release Checklist

This checklist outlines the standard procedure for deploying new versions of the project.

1.  **Deploy to TEST Environment**
    -   From your local machine, run the deployment script for the test environment:
        ```sh
        npm run deploy:test
        ```

2.  **Run In-Sheet Setup (if necessary)**
    -   Open the **TEST** Google Sheet.
    -   If new scopes (permissions) or triggers have been added, you **must** run the setup routine.
    -   Navigate to `🚀 Project Actions > ⚙️ Setup & Configuration > Run Full Setup (Install Triggers)`.
    -   Follow the authorization prompts.

3.  **Perform Smoke Test**
    -   Execute the manual smoke test plan located in `tests/smoke-test.md`.
    -   At a minimum, verify the following core automations:
        -   Progress sync between `Forecasting` and `Upcoming` sheets.
        -   `Permits` status set to `approved` correctly transfers a row to `Upcoming`.
        -   `Delivered` checkbox correctly transfers a row to `Inventory_Elevators`.
        -   `Progress` status set to `In Progress` correctly transfers a row to `Framing`.

4.  **Verify Executions**
    -   Open the Apps Script project and check the **Executions** tab.
    -   Ensure that all recent executions triggered by your smoke test have completed successfully (no red errors).

5.  **Deploy to PRODUCTION Environment**
    -   Once the TEST deployment is verified, deploy to production:
        ```sh
        npm run deploy:prod
        ```

6.  **Run In-Sheet Setup on Production (if necessary)**
    -   Open the **PRODUCTION** Google Sheet.
    -   If you had to run the setup routine in the TEST environment, you **must** also run it here.

7.  **Monitor Production**
    -   After deployment, monitor the **Executions** log for the first few hours or days to ensure no unexpected errors arise from real-world usage.

