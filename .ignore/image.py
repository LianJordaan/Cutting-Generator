import tempfile
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4  # removed landscape
from reportlab.pdfgen import canvas
import os

# ==== 1. Global shape definitions ====

def draw_label_with_relative_and_pixel_offset(ax, base_x, base_y, rel_offset_x, rel_offset_y, pixel_offset_x, pixel_offset_y, label_text, align='center'):
    fig = ax.figure
    fig.canvas.draw()

    # Get current data limits
    x_range = ax.get_xlim()
    y_range = ax.get_ylim()
    width = x_range[1] - x_range[0]
    height = y_range[1] - y_range[0]

    # Apply relative offset in data units
    target_x = base_x + rel_offset_x * width
    target_y = base_y + rel_offset_y * height

    # Convert to display (pixel) coords
    trans = ax.transData
    disp_x, disp_y = trans.transform((target_x, target_y))

    # Apply pixel offset
    disp_x += pixel_offset_x
    disp_y += pixel_offset_y

    # Convert back to data coordinates
    final_x, final_y = trans.inverted().transform((disp_x, disp_y))

    # Draw the label
    ax.text(final_x, final_y, label_text, fontsize=10, color='blue',
            ha=align, va='center')

def value_to_relative(min_val, max_val, value):
    if max_val == min_val:
        raise ValueError("max_val and min_val cannot be the same (division by zero).")
    
    return (value - min_val) / (max_val - min_val)

# ==== 2. Function: build shape from tuple ====
def build_shape_from_tuple(data_tuple):
    shape_type = data_tuple[0]
    label_values = list(data_tuple[1:])  # 5 labels
    label_values[2] = str(label_values[2]) + "x"

    if len(label_values) != 5:
        raise ValueError("Expected 5 label values: length, width, amount, value1, value2")

    length, width, amount, value1, value2 = label_values

    if shape_type == "11":
        template = {
            "lines": [
                (0, 0, length, 0),
                (length, 0, length, width),
                (length, width, length - value1, width),
                (length - value1, width, 0, value2),
                (0, value2, 0, 0),
            ],
            "label_positions": [
                (0.5, 0, 0, -10, 'center'),    # label below center
                (1.0, 0.5, 0, 0, 'left'),        # right side, offset right
                (0.5, 0.5, 0, 0, 'center'),     # middle center
                (value_to_relative(0, length, (length)-(value1/2)), 1.0, 0, 5, 'center'),     # near top right
                (0.0, value_to_relative(0, width, value2/2), -5, 0, 'right'),      # near bottom left
            ],
        }

    elif shape_type == "12":
        template = {
            "lines": [
                (0, 0, length, 0),
                (length, 0, length, width),
                (length, width, length - value1, width),
                (length - value1, width, length - value1, value2),
                (length - value1, value2, 0, value2),
                (0, value2, 0, 0),
            ],
            "label_positions": [
                (0.5, 0, 0, -10, 'center'),    # label below center
                (1.0, 0.5, 0, 0, 'left'),        # right side, offset right
                (0.5, 0.5, 0, 0, 'center'),     # middle center
                (value_to_relative(0, length, (length)-(value1/2)), 1.0, 0, 5, 'center'),     # near top right
                (0.0, value_to_relative(0, width, value2/2), -5, 0, 'right'),      # near bottom left
            ],
        }

    elif shape_type == "21":
        template = {
            "lines": [
                (0, 0, length, 0),                             # bottom
                (length, 0, length, value2),                   # right vertical (lower part)
                (length, value2, value1, width),      # sloped part
                (value1, width, 0, width),            # top
                (0, width, 0, 0),                              # left vertical
            ],
            "label_positions": [
                (0.5, 0.0, 0, -10, 'center'),  # centered below
                (0.0, 0.5, -5, 0, 'right'),  # left mid
                (0.5, 0.5, 0, 0, 'center'),  # center of the shape
                (value_to_relative(0, length, value1 / 2), 1.0, 0, 5, 'center'),  # near top left
                (1.0, value_to_relative(0, width, value2/2), 0, 0, 'left'),      # near bottom right
            ],
        }

    elif shape_type == "22":
        template = {
            "lines": [
                (0, 0, length, 0),                             # bottom
                (length, 0, length, value2),                   # right vertical (lower part)
                (length, value2, value1, value2),               # horizontal part   
                (value1, value2, value1, width),               # sloped part
                (value1, width, 0, width),                     # top
                (0, width, 0, 0),                              # left vertical
            ],
            "label_positions": [
                (0.5, 0.0, 0, -10, 'center'),  # centered below
                (0.0, 0.5, -5, 0, 'right'),  # left mid
                (0.5, 0.5, 0, 0, 'center'),  # center of the shape
                (value_to_relative(0, length, value1 / 2), 1.0, 0, 5, 'center'),  # near top left
                (1.0, value_to_relative(0, width, value2/2), 5, 0, 'left'),      # near bottom right
            ],
        }

    elif shape_type == "31":
        template = {
            "lines": [
                (length, 0, length, width),
                (length, width, 0, width),
                (0, width, 0, width-value2),
                (0, width-value2, length-value1, 0),
                (length-value1, 0, length, 0),
            ],
            "label_positions": [
                (0.5, 1, 0, 5, 'center'),    # label above center
                (1.0, 0.5, 0, 0, 'left'),        # right side, offset right
                (0.5, 0.5, 0, 0, 'center'),     # middle center
                (value_to_relative(0, length, (length)-(value1/2)), 0, 0, -10, 'center'),     # near top right
                (0.0, value_to_relative(0, width, (width) - (value2/2)), -5, 0, 'right'),      # near bottom left
            ],
        }

    elif shape_type == "32":
        template = {
            "lines": [
                (length, 0, length, width),
                (length, width, 0, width),
                (0, width, 0, width-value2),
                (0, width-value2, length-value1, width-value2),
                (length-value1, width-value2, length-value1, 0),
                (length-value1, 0, length, 0),
            ],
            "label_positions": [
                (0.5, 1, 0, 5, 'center'),    # label above center
                (1.0, 0.5, 0, 0, 'left'),        # right side, offset right
                (0.5, 0.5, 0, 0, 'center'),     # middle center
                (value_to_relative(0, length, (length)-(value1/2)), 0, 0, -10, 'center'),     # near bottom right
                (0.0, value_to_relative(0, width, (width) - (value2/2)), -5, 0, 'right'),      # near top left
            ],
        }

        
    elif shape_type == "41":
        template = {
            "lines": [
                (0, 0, value1, 0),
                (value1, 0, length, width-value2),
                (length, width-value2, length, width),
                (length, width, 0, width),
                (0, width, 0, 0),
            ],
            "label_positions": [
                (0.5, 1, 0, 5, 'center'),    # label above center
                (0, 0.5, -5, 0, 'right'),        # right side, offset right
                (0.5, 0.5, 0, 0, 'center'),     # middle center
                (value_to_relative(0, length, value1 / 2), 0, 0, -10, 'center'),  # near top left
                (1.0, value_to_relative(0, width, (width) - (value2/2)), 0, 0, 'left'),      # near bottom right
            ],
        }   

    elif shape_type == "42":
        template = {
            "lines": [
                (0, 0, value1, 0),
                (value1, 0, value1, width-value2),
                (value1, width-value2, length, width-value2),
                (length, width-value2, length, width),
                (length, width, 0, width),
                (0, width, 0, 0),
            ],
            "label_positions": [
                (0.5, 1, 0, 5, 'center'),    # label above center
                (0, 0.5, -5, 0, 'right'),        # right side, offset right
                (0.5, 0.5, 0, 0, 'center'),     # middle center
                (value_to_relative(0, length, value1 / 2), 0, 0, -10, 'center'),  # near top left
                (1.0, value_to_relative(0, width, (width) - (value2/2)), 0, 0, 'left'),      # near bottom right
            ],
        }

    else:
        raise ValueError(f"Unknown shape type: {shape_type}")

    return {
        "type": shape_type,
        "lines": template["lines"],
        "label_positions": template["label_positions"],
        "label_values": label_values,
    }

def draw_shape(ax, shape):
    lines = shape["lines"]
    labels = shape["label_values"]
    label_positions = shape["label_positions"]

    # 1. Draw lines
    for (x1, y1, x2, y2) in lines:
        ax.plot([x1, x2], [y1, y2], 'k')

    # 2. Compute bounding box of shape
    all_x = [x for line in lines for x in (line[0], line[2])]
    all_y = [y for line in lines for y in (line[1], line[3])]
    min_x, max_x = min(all_x), max(all_x)
    min_y, max_y = min(all_y), max(all_y)

    # 3. Set limits before drawing labels so they match
    ax.set_xlim(min_x - 10, max_x + 10)
    ax.set_ylim(min_y - 10, max_y + 10)

    ax.set_aspect('equal')
    # 4. Draw labels
    for i, label_info in enumerate(label_positions):
        if i >= len(labels):
            continue

        if len(label_info) != 5:
            raise ValueError("Each label position must be a 5-tuple: (rel_x, rel_y, offset_x, offset_y, align)")

        rel_x, rel_y, offset_x, offset_y, align = label_info
        label = str(labels[i])

        # Use bounding box min as the anchor
        draw_label_with_relative_and_pixel_offset(
            ax,
            base_x=0,
            base_y=0,
            rel_offset_x=rel_x,
            rel_offset_y=rel_y,
            pixel_offset_x=offset_x,
            pixel_offset_y=offset_y,
            label_text=label,
            align=align
        )

    ax.grid(True)  # Show grid for easier visual alignment

    ax.axis('off')

# ==== 4. Generate one batch image (12 shapes max) ====
def plot_shapes_batch(shapes, batch_num, output_dir):
    fig, axes = plt.subplots(4, 3, figsize=(8.27, 11.69))  # A4 portrait
    axes = axes.flatten()

    for i, shape in enumerate(shapes):
        draw_shape(axes[i], shape)

    for j in range(len(shapes), 12):
        axes[j].axis('off')

    plt.tight_layout()
    filepath = os.path.join(output_dir, f"batch_{batch_num}.png")
    plt.savefig(filepath, dpi=300)
    plt.close()
    return filepath

# ==== 5. Compile images into a PDF ====
def create_pdf_from_images(image_paths, output_path="output.pdf"):
    c = canvas.Canvas(output_path, pagesize=A4)  # portrait
    width, height = A4

    for img_path in image_paths:
        c.drawImage(img_path, 0, 0, width=width, height=height)
        c.showPage()

    c.save()

# ==== 6. Main entry: generate PDF from shape tuples ====
def shapes_to_pdf(shape_tuples, output_pdf="cutout_shapes.pdf"):
    batch_size = 12
    image_paths = []

    shape_objects = [build_shape_from_tuple(t) for t in shape_tuples]

    # Create a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        for i in range(0, len(shape_objects), batch_size):
            batch = shape_objects[i:i+batch_size]
            img_path = plot_shapes_batch(batch, i // batch_size, output_dir=temp_dir)
            image_paths.append(img_path)

        create_pdf_from_images(image_paths, output_pdf)

    print(f"✅ PDF created with {len(image_paths)} pages: {output_pdf}")
    # Temporary directory and its contents are automatically cleaned up


# ==== 7. Example usage ====
if __name__ == "__main__":
    shape_tuples = [
        ("11", 1001, 502, 1, 200, 122),
        ("12", 1003, 504, 1, 560, 244),
        ("21", 1005, 506, 1, 560, 226),
        ("22", 1007, 508, 1, 560, 218),
        ("31", 1009, 510, 1, 540, 250),
        ("32", 1011, 512, 1, 560, 242),
        ("41", 1013, 514, 1, 560, 234),
        ("42", 1015, 516, 1, 560, 226)
    ]

    # Fill up to more than one page
    shape_tuples *= 1  # make ~18 shapes

    shapes_to_pdf(shape_tuples, output_pdf="custom_shapes.pdf")
