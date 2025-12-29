# GitHub Secrets Setup Guide

This guide explains how to set up the required GitHub secret for automatic deployment to Digital Ocean.

## Required Secret

You need to add one secret to your GitHub repository:

### `DIGITALOCEAN_SSH_KEY`

This is the private SSH key that allows GitHub Actions to connect to your Digital Ocean server.

## Setup Steps

### Option 1: Use Your Existing SSH Key (Recommended if you already have access)

If you can already SSH into the server with `ssh root@139.59.8.81`, you can use your existing private key.

1. **Find your existing private key:**
   ```bash
   # Usually located at:
   ~/.ssh/id_rsa
   # or
   ~/.ssh/id_ed25519
   # or
   ~/.ssh/id_ecdsa
   ```

2. **Copy the private key content:**
   ```bash
   cat ~/.ssh/id_rsa
   # or whichever key file you use
   ```

3. **Add to GitHub Secrets:**
   - Go to your GitHub repository (SamayElectroBD)
   - Click **Settings** → **Secrets and variables** → **Actions**
   - Click **New repository secret**
   - Name: `DIGITALOCEAN_SSH_KEY`
   - Value: Paste the entire private key content (including `-----BEGIN` and `-----END` lines)
   - Click **Add secret**

### Option 2: Create a New SSH Key (Recommended for security)

1. **Generate a new SSH key pair:**
   ```bash
   ssh-keygen -t ed25519 -C "github-actions-deploy" -f ~/.ssh/github_actions_deploy
   ```
   - Press Enter to accept default location
   - Optionally set a passphrase (you can leave it empty for automation)

2. **Copy the public key to your Digital Ocean server:**
   ```bash
   ssh-copy-id -i ~/.ssh/github_actions_deploy.pub root@139.59.8.81
   ```
   
   Or manually:
   ```bash
   cat ~/.ssh/github_actions_deploy.pub | ssh root@139.59.8.81 "mkdir -p ~/.ssh && cat >> ~/.ssh/authorized_keys"
   ```

3. **Test the connection:**
   ```bash
   ssh -i ~/.ssh/github_actions_deploy root@139.59.8.81
   ```
   If it works, you can exit with `exit`

4. **Add private key to GitHub Secrets:**
   ```bash
   cat ~/.ssh/github_actions_deploy
   ```
   - Copy the entire output (including `-----BEGIN` and `-----END` lines)
   - Go to your GitHub repository → **Settings** → **Secrets and variables** → **Actions**
   - Click **New repository secret**
   - Name: `DIGITALOCEAN_SSH_KEY`
   - Value: Paste the private key content
   - Click **Add secret**

## How to Add Secret in GitHub

1. Navigate to your repository: `https://github.com/YOUR_USERNAME/SamayElectroBD`
2. Click on **Settings** (top menu)
3. In the left sidebar, click **Secrets and variables** → **Actions**
4. Click **New repository secret** button
5. Fill in:
   - **Name:** `DIGITALOCEAN_SSH_KEY`
   - **Secret:** Paste your private SSH key (the entire content)
6. Click **Add secret**

## Verify Setup

1. Make a small change in your backend code
2. Commit and push to `main` branch
3. Go to **Actions** tab in GitHub
4. You should see the workflow running
5. Check the logs to ensure deployment succeeded

## Security Best Practices

- ✅ Never commit private keys to the repository
- ✅ Use GitHub Secrets for all sensitive data
- ✅ Consider using a dedicated deployment user instead of root
- ✅ Regularly rotate SSH keys
- ✅ Use key-based authentication (not passwords)

## Troubleshooting

### "Permission denied (publickey)" error
- Verify the public key is in `~/.ssh/authorized_keys` on the server
- Check file permissions: `chmod 600 ~/.ssh/authorized_keys` and `chmod 700 ~/.ssh`
- Ensure the private key in GitHub Secrets includes the header/footer lines

### "Host key verification failed"
- The workflow automatically adds the host to known_hosts, but if issues persist, check the server's SSH configuration

### "pm2: command not found"
- Ensure PM2 is installed globally: `npm install -g pm2`
- Or use the full path: `/usr/bin/pm2` or wherever PM2 is installed

