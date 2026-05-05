import requests
import re


def format_address(address):
    """
    Format address string for API calls.
    Converts to title case and removes extra whitespace.
    
    Args:
        address (str): Address string to format
        
    Returns:
        str: Formatted address string
    """
    if not address:
        return ""
    
    # Remove extra whitespace and convert to title case
    formatted = ' '.join(address.split())
    return formatted.title()


def get_cep_from_address(address):
    """
    Get CEP (postal code) from address using ViaCEP API.
    
    Args:
        address (str): Address string to lookup
        
    Returns:
        str: CEP string in format 'XXXXX-XXX' or None if not found/error
    """
    try:
        formatted_address = format_address(address)
        if not formatted_address:
            return None
            
        # ViaCEP API endpoint for address to CEP
        url = f"https://viacep.com.br/ws/{formatted_address}/json/"
        
        response = requests.get(url, timeout=5)
        response.raise_for_status()
        
        data = response.json()
        
        # If API returns an array with one result or a dict
        if isinstance(data, list):
            if len(data) > 0 and 'cep' in data[0]:
                cep = data[0]['cep']
                return cep if cep != '00000-000' else None
        elif isinstance(data, dict) and 'cep' in data:
            cep = data['cep']
            return cep if cep != '00000-000' else None
            
        return None
    except requests.exceptions.RequestException:
        return None
    except (KeyError, IndexError, ValueError):
        return None


def get_address_from_cep(cep):
    """
    Get address information from CEP using ViaCEP API.
    
    Args:
        cep (str): CEP string (with or without hyphen, 8 digits)
        
    Returns:
        dict: Address information or None if not found/error
    """
    try:
        # Clean CEP: remove non-digit characters
        clean_cep = re.sub(r'\D', '', cep)
        
        # Validate CEP format: must be exactly 8 digits
        if len(clean_cep) != 8 or not clean_cep.isdigit():
            return None
            
        # ViaCEP API endpoint for CEP to address
        url = f"https://viacep.com.br/ws/{clean_cep}/json/"
        
        response = requests.get(url, timeout=5)
        response.raise_for_status()
        
        data = response.json()
        
        # Check if CEP was not found
        if 'erro' in data:
            return None
            
        return data
    except requests.exceptions.RequestException:
        return None
    except (KeyError, ValueError):
        return None