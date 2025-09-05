# Production Deployment Guide

This guide details the steps for moving the PowerPoint Accessibility App to a production environment using a cloud VM with Ubuntu Debian.

## Prerequisites

- A VM with Ubuntu Debian
- Firewall open to port 8501
- SSH access to your VM

## Step 1: Connect to Your VM

SSH into your VM:
```bash
ssh username@your-vm-ip
```

## Step 2: Update System and Install Git

Update the package list:
```bash
sudo apt update
```

Install Git:
```bash
sudo apt install git
```

## Step 3: Clone the Repository

Clone the repository:
```bash
git clone <repository-link>
cd pilot  # or whatever your repository folder is named
```

## Step 4: Install Docker Dependencies

Install required packages:
```bash
sudo apt-get install -y \
    ca-certificates \
    curl \
    gnupg \
    lsb-release
```

## Step 5: Add Docker's Official GPG Key

```bash
sudo install -m 0755 -d /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/debian/gpg | \
  sudo gpg --dearmor -o /etc/apt/keyrings/docker.gpg
sudo chmod a+r /etc/apt/keyrings/docker.gpg
```

## Step 6: Set Up Docker Repository

Set up the Debian (bookworm) repository:
```bash
echo \
  "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.gpg] \
  https://download.docker.com/linux/debian \
  bookworm stable" | \
  sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
```

## Step 7: Install Docker

Update package lists:
```bash
sudo apt-get update
```

Install Docker & Docker Compose:
```bash
sudo apt-get install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
```

## Step 8: Verify Installation

Check that Docker is installed correctly:
```bash
docker --version
docker compose version
```

## Step 9: Deploy the Application

Navigate to your project directory:
```bash
cd ~/pilot  # or your repository folder name
```

Set your Google API key in the docker-compose.yml file:
```bash
nano docker-compose.yml
```

Update the `GOOGLE_API_KEY` value with your actual API key.

## Step 10: Start the Application

Build and run the application in detached mode:
```bash
sudo docker compose up --build -d
```

## Step 11: Access Your Application

Your PowerPoint Accessibility App will be available at:
```
http://your-vm-ip:8501
```

## Additional Notes

- The `-d` flag runs the container in detached mode (in the background)
- Use `sudo docker compose logs` to view application logs
- Use `sudo docker compose down` to stop the application
- Make sure your firewall allows traffic on port 8501

## Troubleshooting

If you encounter issues:
1. Check the logs: `sudo docker compose logs`
2. Verify the container is running: `sudo docker ps`
3. Ensure your Google API key is correctly set in docker-compose.yml
4. Check that port 8501 is open in your firewall
