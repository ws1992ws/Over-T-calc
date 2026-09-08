# Overtime Calculator v3.1

Fix:
- Correctly loads saved projects that contain no Vacation/Intedab exclusions.
- Prevents KeyError when the exclusion table is empty.
- Existing save/load, Excel export, and calculation features remain unchanged.

To update:
1. Replace app.py in your GitHub repository with this version.
2. requirements.txt can remain the same, but replacing it is also fine.
3. Commit the change.
4. Streamlit should redeploy automatically.
