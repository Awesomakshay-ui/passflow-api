#!/usr/bin/env bash
set -e
# Install system libraries for HarfBuzz text shaping
apt-get update -qq && apt-get install -y -qq libharfbuzz-dev libfreetype-dev pkg-config
# Install Python dependencies
pip install -r requirements.txt
