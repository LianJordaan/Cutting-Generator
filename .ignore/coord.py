import matplotlib.pyplot as plt

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

# === Example usage ===
fig, ax = plt.subplots(figsize=(5, 5))
ax.grid(True)
ax.set_xlim(0, 500)
ax.set_ylim(0, 500)

# Data point coordinates (red dot)
base_x = 0
base_y = 0  # Bottom of y-axis
ax.plot(50, 50, 'ro')  # Red dot at data point

# Label offset configuration (relative and pixel offsets are ONLY for the label)
rel_offset_x = 0.5       # relative offset for label (in axis fraction)
rel_offset_y = 0       # relative offset for label (in axis fraction)
pixel_offset_x = 0     # pixel offset for label
pixel_offset_y = -10   # pixel offset for label (move label below the axis)
align = 'center'

# Draw the label (offset from the data point)
draw_label_with_relative_and_pixel_offset(
    ax,
    base_x, base_y,
    rel_offset_x, rel_offset_y,
    pixel_offset_x, pixel_offset_y,
    label_text="Centered Below X Axis",
    align=align
)

plt.title("Label offset: right below x-axis")
plt.show()
