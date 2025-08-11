import os
import pandas as pd
from pptx import Presentation
from template_1 import financials_layout_one
from template_2 import financials_layout_two
from path_helpers import get_base_path

BASE_PATH = get_base_path()

df = pd.read_excel(
    os.path.join(BASE_PATH, "database_strips_v15stale.xlsx"),
    sheet_name="Python Financials Mask",
    header=1,  # because the headers for the table are on row 2
    usecols="B:AK"  # adjust based on column range of the table/info
)

prs = Presentation(os.path.join(BASE_PATH, "financials_templates.pptx"))

for i, layout in enumerate(prs.slide_layouts):
    print(f"Layout {i}: {layout.name}")

df = df.dropna(subset=['pb_id'])

print(f"✅ Loaded {len(df)} buyers from mask.")

def run_strips_template(template_number: int, prs: Presentation, df: pd.DataFrame):
    """
    Wrapper function to select the template and populate the presentation
    with all slides needed, slicing the DataFrame into chunks automatically.

    Parameters
    ----------
    template_number : int
        The name of the template to deploy. Examples might include:
        - 1: The primary buyers strip layout with standard table format.
        - 2: (future) An alternative layout.

    prs : pptx.Presentation
        The loaded PowerPoint Presentation object where slides will be added.

    df : pandas.DataFrame
        The full DataFrame of buyers data. This function will handle
        slicing it into chunks (6 rows per slide).

    Raises
    ------
    ValueError
        If an unknown template number is provided.

    Examples
    --------
    >>> run_strips_template(template_number=1, prs=prs, df=df)
    This would add a slide using layout_one to the existing presentation `prs`
    and populate it with the data from `chunk_df`.
    """

    if template_number == 1:
        rows_per_slide = 5
        runs_total = (len(df) + rows_per_slide - 1) // rows_per_slide # Add an extra to force floor division to work like ceiling division, so last partial slide is included

        for run_count in range(runs_total):
            print(f"📊 Creating slide {run_count + 1} of {runs_total}...")
            start_idx = run_count * rows_per_slide
            chunk_df = df.iloc[start_idx : start_idx + rows_per_slide]
            start_number = run_count * rows_per_slide + 1

            financials_layout_one(
                prs, layout_index=1, buyers_chunk_df=chunk_df, start_number=start_number
            )
        print(f"✅ Finished presentation with {runs_total} slides.")
    if template_number == 2:
        rows_per_slide = 5
        runs_total = (len(df) + rows_per_slide - 1) // rows_per_slide # Add an extra to force floor division to work like ceiling division, so last partial slide is included

        for run_count in range(runs_total):
            print(f"📊 Creating slide {run_count + 1} of {runs_total}...")
            start_idx = run_count * rows_per_slide
            chunk_df = df.iloc[start_idx : start_idx + rows_per_slide]
            start_number = run_count * rows_per_slide + 1

            financials_layout_two(
                prs, layout_index=1, buyers_chunk_df=chunk_df, start_number=start_number
            )
        print(f"✅ Finished presentation with {runs_total} slides.")
    
run_strips_template(2, prs=prs, df=df)
prs.save(os.path.join(BASE_PATH, "buyers_presentation.pptx"))