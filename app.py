from pptx import Presentation
from pptx.util import Inches
import os

# Crea una presentazione
prs = Presentation()

# Percorso della cartella con le immagini
images_folder = 'Images/'

# Lista di tutte le immagini nella cartella
images = [img for img in os.listdir(images_folder) if img.endswith(('.png', '.jpg', '.jpeg'))]

# Aggiungi ogni immagine in una nuova slide
for image in images:
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Usa uno slide layout vuoto
    left = top = Inches(1)
    pic = slide.shapes.add_picture(os.path.join(images_folder, image), left, top, width=Inches(5.5))

# Salva la presentazione
prs.save('/mnt/c/Users/Marco/Desktop/your_presentation.pptx')


