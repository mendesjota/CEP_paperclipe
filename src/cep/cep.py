import logging
import random
import requests
import re
import time
from threading import Lock

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('cep')

RateLimitConfig = {
    "max_requests_per_second": 10,
    "max_retries": 3,
    "retry_base_delay": 1.0,
    "retry_max_delay": 10.0,
    "timeout": 5
}

API_ENDPOINTS = {
    "viacep": "https://viacep.com.br/ws/{query}/json/",
    "brasilapi": "https://brasilapi.com.br/api/cep/v1/{query}"
}

_rate_limiter_lock = Lock()
_last_request_time = 0
_min_request_interval = 1.0 / RateLimitConfig["max_requests_per_second"]

CacheConfig = {
    "ttl_seconds": 86400,
    "max_entries": 1000
}

_cep_cache = {}
_cache_lock = Lock()


def _get_from_cache(key):
    with _cache_lock:
        if key in _cep_cache:
            entry = _cep_cache[key]
            if time.time() < entry["expires_at"]:
                logger.debug(f"Cache hit for key: {key}")
                return entry["data"]
            else:
                del _cep_cache[key]
                logger.debug(f"Cache expired for key: {key}")
    return None


def _set_cache(key, data):
    with _cache_lock:
        if len(_cep_cache) >= CacheConfig["max_entries"]:
            oldest_key = min(_cep_cache.keys(), key=lambda k: _cep_cache[k]["expires_at"])
            del _cep_cache[oldest_key]
            logger.debug(f"Cache full, removed oldest entry: {oldest_key}")
        
        _cep_cache[key] = {
            "data": data,
            "expires_at": time.time() + CacheConfig["ttl_seconds"]
        }
        logger.debug(f"Cached data for key: {key}")


def clear_cache():
    """Clear the CEP cache. Useful for testing."""
    global _cep_cache
    with _cache_lock:
        _cep_cache = {}
    logger.info("CEP cache cleared")


def _rate_limit():
    global _last_request_time
    with _rate_limiter_lock:
        current_time = time.time()
        elapsed = current_time - _last_request_time
        if elapsed < _min_request_interval:
            time.sleep(_min_request_interval - elapsed)
        _last_request_time = time.time()


def _retry_request(func, *args, **kwargs):
    last_exception = None
    for attempt in range(RateLimitConfig["max_retries"]):
        try:
            response = func(*args, **kwargs)
            if response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", RateLimitConfig["retry_base_delay"] * 2))
                logger.warning(f"Rate limited (429), waiting {retry_after}s (attempt {attempt + 1}/{RateLimitConfig['max_retries']})")
                time.sleep(retry_after)
                last_exception = Exception("Rate limited (429)")
                continue
            response.raise_for_status()
            logger.debug(f"Request successful on attempt {attempt + 1}")
            return response
        except requests.exceptions.RequestException as e:
            last_exception = e
            logger.warning(f"Request failed (attempt {attempt + 1}/{RateLimitConfig['max_retries']}): {type(e).__name__}: {e}")
            if attempt < RateLimitConfig["max_retries"] - 1:
                base_delay = RateLimitConfig["retry_base_delay"] * (2 ** attempt)
                jitter = random.uniform(0, base_delay * 0.5)
                delay = min(base_delay + jitter, RateLimitConfig["retry_max_delay"])
                logger.info(f"Retrying in {delay:.2f}s...")
                time.sleep(delay)
    logger.error(f"All {RateLimitConfig['max_retries']} retry attempts failed")
    raise last_exception


def _query_brasilapi_cep(cep):
    """
    Query BrasilAPI as fallback for CEP lookup.
    """
    clean_cep = re.sub(r'\D', '', cep)
    if len(clean_cep) != 8 or not clean_cep.isdigit():
        return None

    url = API_ENDPOINTS["brasilapi"].format(query=clean_cep)
    logger.debug(f"Trying BrasilAPI fallback: {url}")
    try:
        _rate_limit()
        response = _retry_request(requests.get, url, timeout=RateLimitConfig["timeout"])
        data = response.json()
        if "errors" in data:
            logger.info(f"CEP {cep} not found in BrasilAPI")
            return None
        logger.info(f"Address found via BrasilAPI fallback for CEP: {cep}")
        return {
            "cep": data.get("cep", ""),
            "logradouro": data.get("street", ""),
            "complemento": data.get("complement", ""),
            "bairro": data.get("neighborhood", ""),
            "localidade": data.get("city", ""),
            "uf": data.get("state", ""),
            "ibge": data.get("ibge", ""),
            "gia": "",
            "ddd": data.get("ddd", "")
        }
    except Exception as e:
        logger.error(f"BrasilAPI fallback failed for CEP {cep}: {type(e).__name__}: {e}")
        return None


def _query_brasilapi_address(address):
    """
    Query BrasilAPI for CEP from address (fallback).
    Note: BrasilAPI doesn't have address->CEP endpoint, return None.
    """
    logger.info("BrasilAPI does not support address->CEP lookup, returning None")
    return None


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
    Uses cache to avoid repeated API calls.

    Args:
        address (str): Address string to lookup

    Returns:
        str: CEP string in format 'XXXXX-XXX' or None if not found/error
    """
    logger.info(f"Looking up CEP for address: {address}")

    formatted_address = format_address(address)
    if not formatted_address:
        logger.warning(f"Empty formatted address for input: {address}")
        return None

    cache_key = f"cep:{formatted_address.lower()}"
    cached_result = _get_from_cache(cache_key)
    if cached_result is not None:
        logger.info(f"Returning cached CEP for address: {address}")
        return cached_result

    try:
        _rate_limit()

        url = API_ENDPOINTS["viacep"].format(query=formatted_address)
        logger.debug(f"Querying ViaCEP: {url}")

        response = _retry_request(requests.get, url, timeout=RateLimitConfig["timeout"])

        data = response.json()

        if isinstance(data, list):
            if len(data) > 0 and 'cep' in data[0]:
                cep = data[0]['cep']
                if cep != '00000-000':
                    logger.info(f"CEP found via ViaCEP for address: {address} -> {cep}")
                    _set_cache(cache_key, cep)
                    return cep
        elif isinstance(data, dict) and 'cep' in data:
            cep = data['cep']
            if cep != '00000-000':
                logger.info(f"CEP found via ViaCEP for address: {address} -> {cep}")
                _set_cache(cache_key, cep)
                return cep

        logger.info(f"No CEP found for address: {address}")
        return None
    except Exception as e:
        logger.error(f"Error looking up CEP for address {address}: {type(e).__name__}: {e}")
        fallback_result = _query_brasilapi_address(address)
        if fallback_result:
            logger.info(f"CEP found via BrasilAPI fallback for address: {address}")
            _set_cache(cache_key, fallback_result)
            return fallback_result
        return None


def get_address_from_cep(cep):
    """
    Get address information from CEP using ViaCEP API with BrasilAPI fallback.
    Uses cache to avoid repeated API calls.
    
    Args:
        cep (str): CEP string (with or without hyphen, 8 digits)
        
    Returns:
        dict: Address information or None if not found/error
    """
    logger.info(f"Looking up address for CEP: {cep}")

    clean_cep = re.sub(r'\D', '', cep)

    if len(clean_cep) != 8 or not clean_cep.isdigit():
        logger.warning(f"Invalid CEP format: {cep}")
        return None

    cache_key = f"address:{clean_cep}"
    cached_result = _get_from_cache(cache_key)
    if cached_result is not None:
        logger.info(f"Returning cached address for CEP: {cep}")
        return cached_result

    try:
        _rate_limit()

        url = API_ENDPOINTS["viacep"].format(query=clean_cep)
        logger.debug(f"Querying ViaCEP: {url}")

        response = _retry_request(requests.get, url, timeout=RateLimitConfig["timeout"])

        data = response.json()

        if 'erro' in data:
            logger.info(f"CEP {cep} not found in ViaCEP, trying BrasilAPI fallback")
            result = _query_brasilapi_cep(cep)
            if result:
                _set_cache(cache_key, result)
            return result

        logger.info(f"Address found via ViaCEP for CEP: {cep}")
        _set_cache(cache_key, data)
        return data
    except Exception as e:
        logger.warning(f"ViaCEP failed for {cep}: {type(e).__name__}: {e}, trying BrasilAPI fallback")
        result = _query_brasilapi_cep(cep)
        if result:
            _set_cache(cache_key, result)
        return result