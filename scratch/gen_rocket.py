from PIL import Image, ImageDraw

def create_color_rocket_png(filepath):
    # Create 64x64 transparent PNG rocket icon
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # 1. Fire Flame (Orange / Yellow)
    draw.polygon([(20, 48), (14, 58), (24, 52), (28, 62), (32, 50)], fill=(255, 100, 0, 255))
    draw.polygon([(22, 49), (18, 55), (25, 52), (27, 58), (30, 50)], fill=(255, 200, 0, 255))

    # 2. Left & Right Fins (Red / Coral)
    draw.polygon([(16, 36), (8, 48), (22, 44)], fill=(239, 68, 68, 255))
    draw.polygon([(48, 36), (56, 48), (42, 44)], fill=(220, 38, 38, 255))

    # 3. Main Rocket Body (White / Gradient Slate)
    draw.ellipse([(20, 10), (44, 46)], fill=(248, 250, 252, 255))
    draw.polygon([(20, 28), (32, 4), (44, 28)], fill=(248, 250, 252, 255))

    # 4. Red Nose Cone Top
    draw.polygon([(26, 16), (32, 4), (38, 16)], fill=(239, 68, 68, 255))

    # 5. Rocket Metallic Band (Indigo)
    draw.rectangle([(22, 34), (42, 38)], fill=(99, 102, 241, 255))

    # 6. Porthole Outer Ring (Dark Gray/Navy) & Inner Glass (Cyan)
    draw.ellipse([(26, 20), (38, 32)], fill=(30, 41, 59, 255))
    draw.ellipse([(28, 22), (36, 30)], fill=(56, 189, 248, 255))
    draw.ellipse([(29, 23), (32, 26)], fill=(255, 255, 255, 200)) # Glass Highlight

    img.save(filepath, "PNG")

create_color_rocket_png("d:\\AI 코딩\\test\\rocket_icon.png")
print("rocket_icon.png created successfully!")
