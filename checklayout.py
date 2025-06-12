from pptx import Presentation

def check_slide_layouts(prs):
    prs = Presentation("templates/corporate_template.pptx")
    
    print(f"Total slide layouts available: {len(prs.slide_layouts)}\n")
    
    for i, layout in enumerate(prs.slide_layouts):
        print(f"Index {i}: {layout.name}")

# Run the check on your corporate template
check_slide_layouts("templates/corporate_template.pptx")
