import tkinter as tk

def _round_rectangle(canvas, x1, y1, x2, y2, radius=25, **kwargs):
    points = [
        x1+radius, y1,
        x2-radius, y1,
        x2, y1,
        x2, y1+radius,
        x2, y2-radius,
        x2, y2,
        x2-radius, y2,
        x1+radius, y2,
        x1, y2,
        x1, y2-radius,
        x1, y1+radius,
        x1, y1
    ]
    return canvas.create_polygon(points, smooth=True, **kwargs)

def round_button(parent, radius, fill, command=None, text="Button", width=150, height=50, x=0, y=0, text_color="white", font=None):
    canvas = tk.Canvas(parent, width=width, height=height, highlightthickness=0)
    canvas.place(x=x, y=y)

    rect = _round_rectangle(canvas, 0, 0, width, height, radius=radius, fill=fill)
    label = canvas.create_text(width//2, height//2, text=text, fill=text_color, font=font)

    if command:
        def on_click(event):
            command()
        canvas.tag_bind(rect, "<Button-1>", on_click)
        canvas.tag_bind(label, "<Button-1>", on_click)

    return canvas
