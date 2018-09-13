from pptx import Presentation

prs = Presentation('adi_template_picture.pptx')

# Loop through layouts and see various elements
for idx, _ in enumerate(prs.slide_layouts):
  slide = prs.slides.add_slide(prs.slide_layouts[idx])

  for shape in slide.placeholders:
    if shape.is_placeholder:
      phf = shape.placeholder_format

      try:
          shape.text = 'Layout: {}, Placeholder:{}, type:{}'.format(idx, phf.idx, shape.name)
      except AttributeError:
        print('{} has no text attribute'.format(phf.type))

      print('{} {}'.format(phf.idx,shape.name))

prs.save('template_out.pptx')


