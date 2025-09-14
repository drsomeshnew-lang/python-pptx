from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Create a presentation object
prs = Presentation()

# Slide 1: Title Slide
slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = "A Prospective Cohort Study to Determine Prognostic Factors in Non-Segmental Vitiligo"
subtitle.text = "Single-Center Tertiary Care Study (2023–2025)\n[Your Name/Department]\n[Your Institution]\n[Presentation Date]"

# Slide 2: Introduction
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Introduction"
content.text = """• Vitiligo affects 0.36% of the general population
• Non-segmental vitiligo (NSV) has a highly variable and unpredictable course
• No existing prognostic scoring system for NSV
• Objective: Identify prognostic factors and develop a scoring system to predict disease behavior and treatment response"""

# Slide 3: Methodology – Study Design
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Methodology – Study Design"
content.text = """• Design: Prospective cohort study
• Duration: July 2023 – June 2025
• Sample Size: 275 patients (after exclusions)
• Inclusion: NSV or mixed vitiligo, all ages, consented patients
• Exclusion: Segmental vitiligo, other hypopigmentary disorders
• Follow-up: 6 and 12 months"""

# Slide 4: Methodology – Data Collection
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Methodology – Data Collection"
content.text = """Variables Recorded:
• Age of onset, disease duration, treatment delay
• Autoimmune history (personal/family)
• Clinical markers: Koebnerization, confetti, trichrome, leukotrichia
• BSA calculation, psychosocial scores (VISS-22, DLQI, F-VIS, C-DLQI)

Outcomes:
• Progression, repigmentation, treatment dependency
• Patient satisfaction, psychosocial impact"""

# Slide 5: Baseline Demographics
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Baseline Demographics"
content.text = """• Mean Age: 29.03 ± 13.90 years
• Mean Age of Onset: 22.59 ± 14.37 years
• Mean Disease Duration: 6.76 ± 6.90 years
• Treatment Delay: 7.76 ± 12.45 months
• Female : Male: 51.1% : 48.9%
• Family History Vitiligo: 13.82%
• Personal Autoimmunity: 8.73%"""

# Slide 6: Collinearity Among Variables
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Collinearity Among Variables"
content.text = """Heatmap showed correlations:
• Koebnerization ↔ Trichrome ↔ Confetti
• Mucosal ↔ Periorificial ↔ Acral
• Family history of vitiligo ↔ Family history of autoimmunity
• Treatment modalities were mutually exclusive"""

# Slide 7: Predictors of Larger BSA (≥5%)
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Predictors of Larger BSA (≥5%)"
content.text = """Significant predictors at 6 & 12 months:
• Leukotrichia (OR: 3.69 → 3.48)
• Confetti lesions (OR: 3.29 → 2.69)
• Trichrome pattern (OR: 3.10 → 2.52)
• Acral lesions (OR: 2.19 → 2.39)
• Koebnerization (OR: 2.10 → 2.03)"""

# Slide 8: Predictors of Progression
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Predictors of Progression"
content.text = """• Family history of autoimmunity significantly predicted progression at 12 months (OR: 3.15, p=0.03)
• Early age of onset (0-15 years) showed negative association
• Confetti, trichrome, Koebnerization did not predict progression—possibly due to aggressive treatment"""

# Slide 9: Predictors of Poor Repigmentation
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Predictors of Poor Repigmentation"
content.text = """• Leukotrichia strongly predicted poor repigmentation:
  - 6 months: OR = 2.50 (p=0.047)
  - 12 months: OR = 3.09 (p=0.018)
• Mucosal, periorificial, acral lesions also showed positive but non-significant trends"""

# Slide 10: Predictors of Treatment Dependency
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Predictors of Treatment Dependency"
content.text = """• Confetti lesions (OR: 2.69-2.93)
• Trichrome pattern (OR: 2.13-2.21)
• Combined treatment modality (OR: 1.50)
• Early age of onset (0-15 years) was protective (OR: 0.49)"""

# Slide 11: Predictors of Psychosocial Impact
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Predictors of Psychosocial Impact"
content.text = """• Age of onset 0-15 years (high impact via F-VIS)
• Koebnerization, periorificial lesions, female sex
• Leukotrichia, mucosal lesions (via DLQI)
• Older age and family history of vitiligo were protective"""

# Slide 12: Discussion – Key Findings
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Discussion – Key Findings"
content.text = """• Leukotrichia, confetti, trichrome, Koebnerization, mucosal/acral involvement are poor prognostic markers
• Family history of autoimmunity predicts progression and treatment dependency
• Early-onset vitiligo has lower extent and dependency but higher psychosocial impact"""

# Slide 13: Clinical Implications
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Clinical Implications"
content.text = """• Incorporate prognostic markers into early risk stratification
• Aggressive combination therapy for high-risk phenotypes
• Psychosocial support for children and young adults
• Periodic re-evaluation of prognostic scores"""

# Slide 14: Limitations
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Limitations"
content.text = """• Single-center, tertiary care setting → selection bias
• Underpowered for multiple outcomes
• No biochemical confirmation of autoimmunity
• Treatment alters natural course"""

# Slide 15: Conclusion
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Conclusion"
content.text = """• First prospective study to identify and validate prognostic factors in NSV
• Proposed scoring system can guide treatment intensity and counselling
• Future directions: Multicenter validation, incorporation of genetic/immunological biomarkers"""

# Slide 16: Acknowledgments
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Acknowledgments"
content.text = """• Patients and families
• Institutional ethics board
• Research team and dermatology department
• Funding sources (if any)"""

# Slide 17: References
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "References"
content.text = """• van Geel et al., J Eur Acad Dermatol Venereol. 2019
• Zhang et al., JAMA Dermatol. 2020
• Kumari et al., Pigment Cell Melanoma Res. 2025
• Mahajan et al., Indian Dermatol Online J. 2019
• ... [Other key references from manuscript]"""

# Slide 18: Q&A
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Q&A"
content.text = "Thank You!\nQuestions?"

# Save the presentation
prs.save('Vitiligo_Prognostic_Study.pptx')

print("PowerPoint presentation created successfully!")
