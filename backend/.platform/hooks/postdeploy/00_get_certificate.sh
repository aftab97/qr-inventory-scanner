#!/usr/bin/env bash
# Place in .platform/hooks/postdeploy directory
sudo certbot -n -d YOURDOMAINHERE --nginx --agree-tos --email YOUREMAILHERE
