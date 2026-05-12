
from PIL import Image, ImageDraw, ImageFont
from brand_assets import load_logo_image

def create_wizard_graphics():
    # 1. WizardImageFile (164x314) - Side banner
    # Pure white with lime accent at bottom and logo at top
    side = Image.new('RGB', (164, 314), (255, 255, 255))
    logo = load_logo_image((120, 120))
    if logo:
        # Center logo horizontally
        x = (164 - logo.width) // 2
        side.paste(logo, (x, 40), logo if logo.mode == 'RGBA' else None)
    
    # Add a lime green bar at the bottom for flair
    draw = ImageDraw.Draw(side)
    draw.rectangle([0, 280, 164, 314], fill='#84cc16')
    
    side.save('wizard_side.bmp')
    print("Created wizard_side.bmp")

    # 2. WizardSmallImageFile (55x55) - Top right small logo
    small = Image.new('RGB', (55, 55), (255, 255, 255))
    logo_s = load_logo_image((45, 45))
    if logo_s:
        x = (55 - logo_s.width) // 2
        y = (55 - logo_s.height) // 2
        small.paste(logo_s, (x, y), logo_s if logo_s.mode == 'RGBA' else None)
    
    small.save('wizard_top.bmp')
    print("Created wizard_top.bmp")

if __name__ == '__main__':
    create_wizard_graphics()
