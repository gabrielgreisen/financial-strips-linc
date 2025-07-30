def clean_domain(raw_domain):
    """
    Removes 'www.' prefix from a domain string if present.

    Parameters
    ----------
    raw_domain : str
        The domain string, e.g. 'www.afya.com.br'.

    Returns
    -------
    str
        The cleaned domain, e.g. 'afya.com.br'.
    """

    if not isinstance(raw_domain, str): # Protects from NaN and None
        return ""
    return raw_domain.replace("www.", "")

def shorten_name(domain):
    """
    Extracts the base name for file naming.
    Example: 
    'afya.com.br' -> 'afya'
    'merative.com' -> 'merative'
    'abc.co.uk' -> 'abc'
    """
    if not isinstance(domain, str):
        return ""
    return domain.split(".")[0]