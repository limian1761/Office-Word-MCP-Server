# syntax=docker/dockerfile:1

# Multi-stage build for different deployment targets

# Stage 1: Base application image (Linux)
FROM python:3.11-slim as linux-base

# Set working directory
WORKDIR /app

# Install build dependencies
RUN apt-get update \
    && apt-get install -y --no-install-recommends build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy project files
COPY . /app

# Install Python dependencies
RUN pip install --no-cache-dir .

# Default command
ENTRYPOINT ["word_docx_tools"]

# Stage 2: Windows container image with COM support
FROM mcr.microsoft.com/windows/servercore:ltsc2022 as windows-base

# Install Python
# Note: This is a simplified example. In practice, you would need to install Python and dependencies
# Windows containers with full COM support for Office applications have additional requirements
# This is provided as a template for when you need Windows container support

# Install Chocolatey for package management
RUN powershell -Command \
    Set-ExecutionPolicy Bypass -Scope Process -Force; \
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; \
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

# Install Python
RUN choco install python311 -y

# Add Python to PATH
ENV PATH="C:\Python311;C:\Python311\Scripts;%PATH%"

# Set working directory
WORKDIR /app

# Copy project files
COPY . /app

# Install Python dependencies
RUN pip install --no-cache-dir .

# Default command
ENTRYPOINT ["word_docx_tools"]