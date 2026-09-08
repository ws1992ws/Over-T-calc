# Overtime Calculator v4

This version rebuilds the project save/load workflow.

Major improvements:
- Uploading a JSON file no longer changes the app automatically.
- User must explicitly click Load Project.
- Period, salary history, transportation history, exclusions, and rules load together.
- Data editors are rebuilt after loading, avoiding stale Streamlit widget state.
- Save Project is beside Calculate Overtime.
- Inputs / Results / How It Works tabs.
- New / Clear Project command.
- Excel report export remains available.

To update your existing Streamlit app:
1. Replace app.py in your GitHub repository with the v4 app.py.
2. Replace requirements.txt if desired (same dependencies).
3. Commit the change.
4. Wait for Streamlit to redeploy.
5. Refresh your existing app URL.

Your existing saved JSON files from v3 should remain compatible.
