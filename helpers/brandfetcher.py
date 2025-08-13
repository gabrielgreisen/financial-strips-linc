import os
import requests
import pandas as pd
from helpers.request_helpers import clean_domain, shorten_name

#BRANDFETCH_API_KEY = "KGTJrCtYxuDa0XfuQF4m1RNHkaC0D+x2tW9KLbsiVp4="

def get_brandfetch_logo(domain, BRANDFETCH_API_KEY = "KGTJrCtYxuDa0XfuQF4m1RNHkaC0D+x2tW9KLbsiVp4="):
    
    """
    Queries the Brandfetch API for a given domain and returns 
    the direct URL to a PNG logo, prefering 'dark' logos over 'light'.

    This function sends an HTTP GET request to Brandfetch's brand endpoint 
    for the specified domain (e.g., 'nestle.com') using your API key. 
    It parses the JSON response and looks for the first available 
    PNG format of a logo lockup (type 'logo').

    Parameters
    ----------
    domain : str
        The domain name to look up, such as 'nestle.com' or 'afya.com.br'.

    Returns
    -------
    str or None
        A direct URL to the PNG logo file if found, else None if the API 
        does not return a logo or an error occurs.

    Example
    -------
    >>> get_brandfetch_logo("nestle.com")
    'https://cdn.brandfetch.io/.../logo.png'
    """

    url = f"https://api.brandfetch.io/v2/brands/{domain}"
    headers = {
        "Authorization" : f"Bearer {BRANDFETCH_API_KEY}"
    }
    
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print(f"Brandfetch did not find a logo for {domain}")
        return None
    
    data = response.json()
    logos = data.get("logos", [])

    # Fisrt look expicitly for a dark theme logo PNG
    for asset in logos:
        if asset.get("type") == "logo" and asset.get("theme") == "dark":
            for fmt in asset.get("formats", []):
                if fmt.get("format") == "png":
                    return fmt.get("src")
                
    # If dark theme is not available, fallback to any logo PNG
    for asset in logos:
        if asset.get("type") == "logo":
            for fmt in asset.get("formats", []):
                if fmt.get("format") == "png":
                    return fmt.get("src")            
    
    return None

def download_logo_file(logo_url, filename, save_dir= "logos"):
    
    """
    Downloads a PNG logo from a given URL and saves it to the local logos folder.

    This function makes an HTTP GET request to download the binary PNG 
    contents from the provided `logo_url`. It saves the file under the 
    specified `save_dir` using the domain name as the filename.

    Parameters
    ----------
    logo_url : str
        The direct URL to the PNG logo file to download.

    filename : str
        The cleaned domain name, used to name the saved PNG file 
        (e.g., 'nestle.com' results in 'logos/nestle.com.png').

    save_dir : str, optional
        The directory where logos should be saved. Defaults to 'logos'.
        If the directory does not exist, it will be created.

    Returns
    -------
    str or None
        The full path to the saved PNG file if the download succeeds, 
        else None if the download fails.

    Example
    -------
    >>> download_logo_file('https://cdn.brandfetch.io/.../logo.png', 'nestle.com')
    'logos/nestle.com.png'
    """

    os.makedirs(save_dir, exist_ok=True)
    path = os.path.join(save_dir, f"{filename}.png")
    r = requests.get(logo_url)
    if r.status_code == 200:
        with open(path, "wb") as f:
            f.write(r.content)
            print(f"✅ Saved logo as {path}")
    else:
        print(f"❌ Failed to download logo from {logo_url}")



if __name__ == "__main__":
    # Load excel data
    df = pd.read_excel(
        "database_strips_logo_download_v2.xlsx",
        sheet_name="API removal",
        header=1,
        usecols="B:K"
    )

    # Clean Domains after loading
    df["clean_domains"] = df['website'].dropna()

    #Loop over unique domains
    for domain in df["clean_domains"].dropna().unique():
        short_name = shorten_name(domain)
        local_path = f"logos/{short_name}.png"
        if os.path.exists(local_path):
            print(f"✅ Logo already exists for {domain}, skipping.")
            continue

        logo_url = get_brandfetch_logo(domain)
        if logo_url:
            download_logo_file(logo_url, short_name)

