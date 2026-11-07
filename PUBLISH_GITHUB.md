# Publishing the clean version to GitHub

1. Create an EMPTY repository on GitHub, for example DailyWorkReports-Flexible.
2. Extract the GitHub source ZIP or use the clean public source folder.
3. Do not import or copy the original private v15.3 / v15.4 working trees.
4. Run public verification: python scripts/verify_public.py
5. Review the contents with git status and git diff --cached before pushing.
6. MIT is adopted for original ProTech-owned code; include LICENSE in any source distribution and preserve third-party notices.
7. Upload Windows EXE and Android test APK as optional GitHub Releases assets;
   do not put compiled builds in the Git source tree.
8. Android Debug APK is for testing; build and sign Release separately.

Example commands from the root of a new local repository:

    git init
    git add .
    git diff --cached --name-only
    git commit -m "Initial flexible public edition 15.5.0"
    git branch -M main
    git remote add origin https://github.com/YOUR_USERNAME/DailyWorkReports-Flexible.git
    git push -u origin main

Do NOT blindly use git add --force on ignored files.
No files are uploaded automatically on your behalf.
