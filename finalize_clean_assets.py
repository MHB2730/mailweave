
import os
from PIL import Image

def finalize_assets():
    logo_src = r'C:\Users\bremn\.gemini\antigravity\brain\deb524f2-7c13-457a-b9eb-587c945e54b5\mailweave_hero_logo_clean_1778066882393.png'
    mark_src = r'C:\Users\bremn\.gemini\antigravity\brain\deb524f2-7c13-457a-b9eb-587c945e54b5\mailweave_mark_clean_1778066939653.png'
    
    target_dir = r'c:\Users\bremn\Documents\MailWeave'
    
    # 1. Save Logo
    img_logo = Image.open(logo_src).convert('RGBA')
    img_logo.save(os.path.join(target_dir, 'mailweave_logo.png'))
    print("Saved clean mailweave_logo.png")
    
    # 2. Save ICO
    img_mark = Image.open(mark_src).convert('RGBA')
    img_mark.thumbnail((256, 256), Image.LANCZOS)
    
    # Create square canvas for ICO
    ico_canvas = Image.new('RGBA', (256, 256), (255, 255, 255, 0))
    mx = (256 - img_mark.width) // 2
    my = (256 - img_mark.height) // 2
    ico_canvas.paste(img_mark, (mx, my), img_mark)
    
    ico_canvas.save(os.path.join(target_dir, 'mailweave.ico'), sizes=[(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)])
    print("Saved clean mailweave.ico")

if __name__ == '__main__':
    finalize_assets()
