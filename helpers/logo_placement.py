from pptx import slide, table

def place_logo_on_slide(slide, table_shape, table, row_idx, col_idx, logo_file, width_spacing, height_spacing, left_spacing, top_spacing):
    """
    Places the logo image on the slide, anchored visually aligned to the table cell
    at (row_idx, col_idx) with a slight downward offset for balanced look.

    Parameters
    ----------
    slide : pptx.slide.Slide
    table_shape : pptx.shape.Shape
        The shape that contains the table, used to get .left and .top
    table : pptx.table.Table
        The actual table object
    row_idx : int
    col_idx : int
    logo_file : str
        Path to the PNG file to insert.
    width_spacing : float
    height_spacing : float
    left_spacing : float
    top_spacing : float
    """
    left = table_shape.left + sum([table.columns[i].width for i in range(col_idx)])
    top = table_shape.top + sum([table.rows[j].height for j in range(row_idx)])
    width = table.columns[col_idx].width
    height = table.rows[row_idx].height

    # Slight inset and slight downward offset to look good even with multi-line text
    img_width = int(width * width_spacing)
    img_height = int(height * height_spacing)
    img_left = left + int(width * left_spacing)
    img_top = top + int(height * top_spacing)  # slightly lower for balanced look

    slide.shapes.add_picture(logo_file, img_left, img_top, img_width, img_height)
