# Deployment Guide

This guide explains how to deploy the Capstone Defense Scheduler to Streamlit Cloud.

## Prerequisites

- A GitHub account
- A Streamlit Cloud account (free at [share.streamlit.io](https://share.streamlit.io))

## Step 1: Prepare Your Repository

1. **Initialize Git (if not already done):**
   ```bash
   git init
   ```

2. **Create .gitignore** (already included in the project)

3. **Add all files:**
   ```bash
   git add .
   ```

4. **Commit:**
   ```bash
   git commit -m "Initial commit: Capstone Defense Scheduler"
   ```

## Step 2: Push to GitHub

1. **Create a new repository on GitHub:**
   - Go to [github.com](https://github.com)
   - Click "New repository"
   - Name it (e.g., "capstone-defense-scheduler")
   - Choose public or private
   - Don't initialize with README (we already have one)

2. **Connect and push:**
   ```bash
   git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
   git branch -M main
   git push -u origin main
   ```

## Step 3: Deploy to Streamlit Cloud

1. **Go to Streamlit Cloud:**
   - Visit [share.streamlit.io](https://share.streamlit.io)
   - Sign in with your GitHub account

2. **Create New App:**
   - Click "New app"
   - Select your repository
   - Select branch: `main`
   - Main file path: `app.py`
   - App URL: (auto-generated or customize)

3. **Deploy:**
   - Click "Deploy"
   - Wait for deployment (usually 1-2 minutes)

4. **Your app is live!**
   - Access at: `https://YOUR_APP_NAME.streamlit.app`
   - Share the URL with others

## Step 4: Update Your App

Whenever you push changes to GitHub:

1. **Make changes locally**
2. **Commit and push:**
   ```bash
   git add .
   git commit -m "Description of changes"
   git push
   ```
3. **Streamlit Cloud automatically redeploys** (usually within 1-2 minutes)

## Configuration

The app uses `.streamlit/config.toml` for configuration. This is automatically used by Streamlit Cloud.

## Troubleshooting

### App Won't Deploy

- Check that `requirements.txt` includes all dependencies
- Ensure `app.py` is in the root directory
- Check the deployment logs in Streamlit Cloud dashboard

### Dependencies Issues

- Make sure all packages in `requirements.txt` are available on PyPI
- Check version compatibility
- Streamlit Cloud uses Python 3.11 by default

### File Size Limits

- Streamlit Cloud has file size limits
- Large Excel files (>100MB) may cause issues
- Consider using example generators instead of committing large files

## Advanced: Custom Domain

Streamlit Cloud supports custom domains:

1. In your app settings, go to "Settings" → "Custom domain"
2. Follow the instructions to configure your domain

## Monitoring

- View app metrics in Streamlit Cloud dashboard
- Check logs for errors
- Monitor usage statistics

---

**Note:** The free tier of Streamlit Cloud is perfect for this application. For production use with high traffic, consider upgrading to a paid plan.

