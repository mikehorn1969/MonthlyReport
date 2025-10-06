# Azure Key Vault access

import os
from typing import Optional
from azure.identity import AzureCliCredential
from azure.keyvault.secrets import SecretClient
import dotenv

dotenv.load_dotenv()  # Load environment variables from .env file if present


def get_azure_credential():
    """Get Azure credential using Azure CLI"""    
    try:
        credential = AzureCliCredential()
        # Test the credential
        credential.get_token("https://vault.azure.net/.default")
        print("Using Azure CLI authentication")
        return credential
    except Exception:
        pass
    

def get_kv_client() -> Optional[SecretClient]:
    """Return a SecretClient if KEY_VAULT_NAME is configured, else None."""
    kv_name = os.environ.get("KEY_VAULT_NAME")
    if not kv_name:
        print("KEY_VAULT_NAME not configured, skipping Key Vault access")
        return None
    
    credential = get_azure_credential()
    if not credential:
        print("No valid Azure authentication found, skipping Key Vault access")
        return None
        
    vault_uri = f"https://{kv_name}.vault.azure.net"
    return SecretClient(vault_url=vault_uri, credential=credential)


def get_secret(env_name: str, kv_secret_name: Optional[str] = None, default_value: Optional[str] = None) -> str:
    """
    Fetch a secret from environment, or (if not set) from Azure Key Vault.
    kv_secret_name defaults to env_name if not supplied.
    default_value is returned if secret is not found anywhere.
    """
    # First try environment variable
    val = os.environ.get(env_name)
    if val:
        return val

    # Try Key Vault if configured and authenticated
    client = get_kv_client()
    if client:
        secret_name = kv_secret_name or env_name
        try:
            result = client.get_secret(secret_name).value 
            return result if result is not None else ""
        except Exception as exc:
            print(f"Failed to retrieve '{secret_name}' from Azure Key Vault: {exc}")

    # Return default value or raise error
    if default_value is not None:
        print(f"Using default value for '{env_name}'")
        return default_value
        
    raise KeyError(
        f"Required secret '{env_name}' not found in environment variables, "
        f"Key Vault not configured/accessible, and no default value provided. "
        f"Please set the '{env_name}' environment variable or configure Azure authentication."
    )

