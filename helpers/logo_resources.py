import os
from helpers.brandfetcher import get_brandfetch_logo, download_logo_file
from path_helpers import get_base_path

BASE_PATH = get_base_path()

def ensure_logo_available(logo_name, domain, logo_base_dir="logos"):
    """
    Checks if the logo PNG exists in the local logos folder.
    If not, attempts to fetch it from Brandfetch using the domain
    and saves it under logo_name.

    Parameters
    ----------
    logo_name : str
        The cleaned name to use for the local PNG file (without .png).

    domain : str
        The company website domain (like 'afya.com.br') to use for Brandfetch.

    logo_base_dir : str
        Directory where logo PNG files are stored.

    Returns
    -------
    str or None
        The full path to the logo file if it exists or was fetched successfully, else None.
    """

    logo_dir = os.path.join(BASE_PATH, logo_base_dir)
    logo_file = os.path.join(logo_dir, f"{logo_name}.png")
    if os.path.exists(logo_file):
        return logo_file
    else:
        print(f"🔍 Attempting to fetch logo for {domain} to save as {logo_name}")
        logo_url = get_brandfetch_logo(domain)
        if logo_url:
            download_logo_file(logo_url, logo_name, save_dir=logo_dir)
            if os.path.exists(logo_file):
                print(f"✅ Successfully fetched and saved logo for {domain}")
                return logo_file
        print(f"❌ Could not obtain logo for {domain}")
        return None
    

def get_logo_file_path_main(row, logo_base_dir="logos"):
    """
    Given a DataFrame row with expected columns for logo name and domain,
    returns the full path to the logo file, ensuring it exists (or fetched).

    Parameters
    ----------
    row : pd.Series
        A row from your DataFrame with logo and domain info.

    logo_base_dir : str
        The directory where logo PNG files are stored.

    Returns
    -------
    str or None
        Full path to the logo PNG file, or None if unavailable.
    """

    logo_name = str(row["logo_file"]) # your 8th column for cleaned logo name
    domain = str(row["website"]) # your 7th column for website domain

    return ensure_logo_available(logo_name, domain, logo_base_dir=logo_base_dir)

def get_logo_file_path(row, logo_name_column, domain_column, logo_base_dir="logos"):
    """
    Given a DataFrame row with expected columns for logo name and domain,
    returns the full path to the logo file, ensuring it exists (or fetched).

    Parameters
    ----------
    row : pd.Series
        A row from your DataFrame with logo and domain info.

    logo_base_dir : str
        The directory where logo PNG files are stored.

    Returns
    -------
    str or None
        Full path to the logo PNG file, or None if unavailable.
    """

    logo_name = str(row[logo_name_column]) 
    domain = str(row[domain_column]) 

    return ensure_logo_available(logo_name, domain, logo_base_dir=logo_base_dir)


def get_lincoln_file_path(logo_name, logo_base_dir="linc_logos"):
    """
    Given a DataFrame row with expected columns for logo name and domain,
    returns the full path to the logo file, ensuring it exists (or fetched).

    Parameters
    ----------
    row : pd.Series
        A row from your DataFrame with logo and domain info.

    logo_base_dir : str
        The directory where logo PNG files are stored.

    Returns
    -------
    str or None
        Full path to the logo PNG file, or None if unavailable.
    """

    domain = "www.lincolninternational.com" # your 7th column for website domain

    return ensure_logo_available(logo_name, domain, logo_base_dir=logo_base_dir)