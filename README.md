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

#### average

* check if presentation contains 3 slides, right aspect ratio, horizontal orientation, one typeface, photos giving by
 task(using compare images with some average), layout check
 
#### first slide

* check if first slide has title, subtitle, correct font sizes


# Plans

* collisions between presentation elements
* add other layouts
* second slide
    * move check layout to proper method
    * check value of text blocks and images
    * font sizes check
* third slide same as second
* add generate recommendations what grade presentation need
* develop proper structure