"""
Base Authentication Classes

This module provides the base classes for Meraki authentication methods.
"""

import os
from dotenv import load_dotenv
from authlib.integrations.requests_client import OAuth2Session

class APIKeyAuth:
    """Standard Key-based authentication"""
    def __init__(self) -> None:
        load_dotenv()
        self.api_key = os.getenv('MERAKI_API_KEY', '')
        if not self.api_key:
            raise ValueError("MERAKI_API_KEY not found in environment variables or .env file.")

    def get_auth_token(self) -> str:
        """Get authentication token for API requests."""
        return  self.api_key

class OAuthAuth:
    """OAuth authentication using manual code input."""
    def __init__(self, client_id=None, client_secret=None, redirect_uri=None):
        load_dotenv()
        self.client_id = client_id or os.getenv('MERAKI_CLIENT_ID', '')
        self.client_secret = client_secret or os.getenv('MERAKI_CLIENT_SECRET', '')
        self.redirect_uri = redirect_uri or os.getenv('MERAKI_REDIRECT_URI', 'https://127.0.0.1:8443/oauth_callback')
        if not self.client_id or not self.client_secret:
            raise ValueError("MERAKI_CLIENT_ID and MERAKI_CLIENT_SECRET must be set in environment or passed to constructor.")
        self.token = None
        self.oauth_session = OAuth2Session(
            client_id=self.client_id,
            client_secret=self.client_secret,
            redirect_uri=self.redirect_uri,
            scope='sdwan:config:read sdwan:config:write'
        )
        self.authorization_endpoint = 'https://as.meraki.com/oauth/authorize'
        self.token_endpoint = 'https://as.meraki.com/oauth/token'

    def get_auth_token(self) -> str:
        """Get authentication token for API requests via OAuth flow."""
        if self.token and 'access_token' in self.token:
            return self.token['access_token']
        # Start OAuth flow
        auth_url, _ = self.oauth_session.create_authorization_url(self.authorization_endpoint)
        print("\n🔐 OAuth Authorization Required")
        print(f"Open this URL in your browser and authorize the app:")
        print(auth_url)
        print("\nAfter authorizing, copy the code from the response and paste it below.")
        code = input("Enter the authorization code: ").strip()
        if not code:
            raise ValueError("No authorization code provided.")
        self.token = self.oauth_session.fetch_token(
            self.token_endpoint,
            code=code
        )
        return self.token['access_token']

