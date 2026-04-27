from Parser import ExcelToPptx

ExcelToPptx.create_presentation(
    data='NPF_parsing_excel2pptx/data.xlsx',
    groups_position=(2, 12),
    funds_position=(16, 29),
    min_bar_ratio=0.6
)
