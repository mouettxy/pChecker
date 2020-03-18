rewrite program to use Office API instead of python-pptx module

# Current features

### Images

* generating screenshots of each slide
* generating "skeleton" of each slide (it contains position of each element)
* compare images on presentation with given images using aHash algorithm
* generating view for layout (described later)

### Layouts

* added first layout

### Analyze

* check if presentation contains 3 slides, right aspect ratio, horizontal orientation, one typeface, photos giving by
 task(using compare images with some average), layout check
* check if slide 1, 2, 3 have right count of images, text blocks, has title, elements not overlaps each other
* added experimental grade to presentation
* realised "detail" analyze
* 
 

# Plans

* add distorted images check
* add other layouts
* develop proper structure