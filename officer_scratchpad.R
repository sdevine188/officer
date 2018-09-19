library(officer)
library(magick)
library(dplyr)
library(stringr)
library(tidyr)
library(purrr)
library(rlang)
library(ggplot2)

# https://davidgohel.github.io/officer/
# https://davidgohel.github.io/officer/articles/officer_reader.html

setwd("C:/Users/Stephen/Desktop/R/officer")
list.files()

# read powerpoint
test <- read_pptx(path = "slidex_test.pptx")
test
str(test)
test$table_styles
test$presentation
print(test)

# get powerpoint summary as a dataframe
test_summary <- pptx_summary(test)
test_summary
str(test_summary)
glimpse(test_summary)

# get bullets from slide 4
test_summary %>% filter(content_type == "paragraph", slide_id == 3)

# get image from slide 4
test_summary %>% filter(content_type == "image", slide_id == 4)
slide4_image_path <- test_summary %>% filter(content_type == "image", slide_id == 4) %>% select(media_file)
slide4_image_path
media_extract(x = test, path = slide4_image_path, target = "slide4_image.png")
slide4_image <- image_read("slide4_image.png")
slide4_image

# get pptx table from slide 6
table1 <- test_summary %>% filter(slide_id == 5, id == 5, content_type == "table cell")
table1

table1 <- table1 %>% select(text, row_id, cell_id) %>% spread(key = cell_id, value = text) %>% select(-row_id)
names(table1) <- table1 %>% slice(1)        
table1 <- table1 %>% filter(row_number() != 1)
glimpse(table1)
table1

# add a slide
layout_summary(test)
test2 <- test

length(test)
length(test2)
# note that even though i only add slide to test2, it also adds slide to test?!

image_file <- "slide4_image.png"
plot <- ggplot(data = iris ) +
        geom_point(mapping = aes(Sepal.Length, Petal.Length), size = 3) +
        theme_minimal()

test2 <- test2 %>% add_slide(layout = "Title and Content", master = "Office Theme") %>%
        ph_with_text(type = "body", str = "A first text", index = 1) %>%
        ph_with_text(type = "title", str = "A title") %>%
        add_slide(layout = "Title and Content", master = "Office Theme") %>%
        ph_with_img(type = "body", index = 1, src = image_file, height = 4, width = 4) %>%
        add_slide(layout = "Title and Content", master = "Office Theme") %>%
        ph_with_gg(type = "body", value = plot, index = 1) %>%
        add_slide(layout = "Title and Content", master = "Office Theme") %>%
        ph_with_table(type = "body", value = head(mtcars) )

length(test)
length(test2)
pptx_summary(test) %>% filter(slide_id > 10)
pptx_summary(test2) %>% filter(slide_id > 10)

print(test2, target = "test2.pptx")



#######################################################################################
#######################################################################################
#######################################################################################


# import word document

# get example doc from package 
example_docx <- system.file(package = "officer", "doc_examples/example.docx")
example_docx

# read in docx
doc <- read_docx(example_docx)
doc

# get doc summary as a data.frame
doc_summary <- docx_summary(doc)
doc_summary
head(doc_summary)

# explore contents by getting count of each content_type
doc_summary %>% count(content_type)

# get paragraphs
doc_summary %>% filter(content_type == "paragraph") %>% head()

# get tables
# Cells positions and values are dispatched in columns row_id, cell_id, text and is_header
doc_summary %>% filter(content_type == "table cell") %>% head()
doc_summary %>% filter(content_type == "table cell") %>% select(row_id, cell_id, text, is_header) %>% head()
doc_summary %>% filter(content_type == "table cell") %>% select(row_id, cell_id, text, is_header) %>% 
        arrange(row_id, cell_id) %>% head()

# create function to reshape stacked table output to get original table
stacked_table <- doc_summary %>% filter(content_type == "table cell") %>% select(row_id, cell_id, text, is_header) %>%
        arrange(row_id, cell_id)
stacked_table

reshape_stacked_table_to_original_table <- function(stacked_table) {
        
        # call create_original_table_rows to loop through rows
        original_table <- map_dfr(.x = stacked_table$row_id, 
                                  .f = ~ create_original_table_rows(row_id_value = .x, stacked_table = stacked_table))
        return(original_table)
}


# create function to loop through stacked table and create original table rows
create_original_table_rows <- function(row_id_value, stacked_table) {
        
        # get current_row from stacked_table
        current_row <- stacked_table %>% filter(row_id == row_id_value)
        
        # spread current_row 
        current_row <- current_row %>% as.tibble() %>% mutate(columns = str_c("column_", cell_id)) %>% 
                select(-c(cell_id, is_header)) %>% spread(key = columns, value = text)
        return(current_row)
}

# test
row_id_value <- 2

# call reshape_stacked_table_to_original_table function
# note that this is not a perfect recreation due to the merged cells in the stacked_table
reshape_stacked_table_to_original_table(stacked_table)


###########################################################################


# create word document
my_doc <- read_docx() 
styles_info(my_doc)

# assign image to add to doc
image_file <- "slide4_image.png"

# create a ggplot to add to doc as a saved png with body_add_img
plot <- ggplot(data = iris ) +
        geom_point(mapping = aes(Sepal.Length, Petal.Length), size = 3) +
        theme_minimal()
plot

# ggplots are normally added directly as argument to body_add_gg to be evaluated
# if you have it saved, you can also add is as a png with body_add_img though
# note that pdf did not render in word doc when added - use png instead
ggsave(filename = "plot_for_word_doc.png", plot = plot)

# create ggplot to add directly as a ggplot to be evaluated
plot2 <- ggplot(data = starwars, aes(x = mass)) + geom_histogram()
plot2

# create image of r logo for slip_in_img
r_logo_img <- file.path( R.home("doc"), "html", "logo.jpg" )
r_logo_img

# view list of styles available for adding new elements
read_docx() %>% styles_info()

# add elements to word doc
my_doc <- my_doc %>% 
        body_add_par(value = "png example using plot w heading 1", style = "heading 1") %>% 
        body_add_img(src = "plot_for_word_doc.png", width = 5, height = 6, style = "centered") %>% 
        body_add_break() %>%
        body_add_par(value = "png example using picture w heading 2", style = "heading 2") %>% 
        body_add_img(src = image_file, width = 5, height = 6, style = "centered") %>% 
        body_add_break() %>%
        body_add_par(value = "ggplot2 example w heading 3", style = "heading 3") %>% 
        body_add_gg(value = plot2, width = 5, height = 6, style = "centered") %>% 
        body_add_break() %>%
        body_add_par("Hello world! (style = Normal)", style = "Normal") %>% 
        body_add_par("", style = "Normal") %>% # blank paragraph
        body_add_table(iris %>% head(), style = "table_template") %>%
        # demonstrate slip_in_* feature to put content at beginning or end of element 
        # this is mainly used to take advantage of word's seqfield numbering functionality
        body_add_par("R logo: ", style = "Normal") %>%
        slip_in_img(src = r_logo_img, style = "strong", 
                    width = .3, height = .3, pos = "after") %>% 
        slip_in_text(" - This is ", style = "strong", pos = "before") %>% 
        slip_in_seqfield(str = "SEQ Figure \u005C* ARABIC",
                         style = 'strong', pos = "before") 
my_doc

# get summary
my_doc_summary <- docx_summary(my_doc)
my_doc_summary %>% head()

# write doc to file
print(my_doc, target = "first_example.docx")


#############################################################


# example of using cursor functions
read_docx() %>%
        body_add_par("paragraph 1", style = "Normal") %>%
        body_add_par("paragraph 2", style = "Normal") %>%
        body_add_par("paragraph 3", style = "Normal") %>%
        body_add_par("paragraph 4", style = "Normal") %>%
        body_add_par("paragraph 5", style = "Normal") %>%
        body_add_par("paragraph 6", style = "Normal") %>%
        body_add_par("paragraph 7", style = "Normal") %>%
        print(target = "cursor_example.docx" )

# then read cursor_example and manipulate with cursor functions
cursor_example_edited <- read_docx(path = "cursor_example.docx") %>%
        
        # default template contains only an empty paragraph
        # Using cursor_begin and body_remove, we can delete it
        cursor_begin() %>% body_remove() %>%
        
        # Let add text at the beginning of the
        # paragraph containing text "paragraph 4"
        cursor_reach(keyword = "paragraph 4") %>%
        slip_in_text("This is ", pos = "before", style = "Default Paragraph Font") %>%
        # add break line before paragraph 4 using slip_in_column_break, which is a thinner break than body_add_break()
        slip_in_column_break(pos = "before") %>%

        # move the cursor forward and end a section
        cursor_forward() %>%
        # add break after paragraph 5
        body_add_break() %>%
        body_add_par("The section stop here", style = "Normal") %>%
        # i don't understand sections really, they don't seem that useful for my purposes
        body_end_section_landscape() %>%
        
        # find next example of "paragraph" text (should be paragraph 6)
        cursor_reach(keyword = "paragraph") %>%
        slip_in_text("--this text should precede paragraph 6, but instead matches first instance of 'paragraph' in doc--", 
                     pos = "before", style = "Default Paragraph Font") %>%
        
        # find paragraph 6 using regex - for some reason, regex doesn't work, although docs say it does??
        cursor_reach(keyword = "grap[h|z] 6") %>%
        # cursor_reach(keyword = ".*6") %>%
        slip_in_text("--found paragraph 6 with cursor_reach regex-- ", pos = "before",
                     style = "Default Paragraph Font") %>%
        
        # move the cursor at the end of the document
        cursor_end() %>%
        body_add_par("The document ends now", style = "Normal") %>%

        # move cursor backward
        cursor_backward() %>%
        cursor_backward() %>%
        slip_in_text("-- moved cursor backward twice from 'this document ends now' element to this element --")
        

print(cursor_example_edited, target = "cursor_example_edited.docx")


#########################################################


# replace or remove content
str1 <- "this is a repeated test paragraph" %>% 
        rep(3) %>% str_c(., collapse = ". ")
str2 <- "Drop that text" 
str3 <- "test para 2" %>% 
        rep(3) %>% paste(collapse = "")

replace_and_remove_doc <- read_docx()  %>% 
        body_add_par(value = str1, style = "Normal") %>% 
        body_add_par(value = str2, style = "centered") %>% 
        body_add_par(value = str3, style = "Normal") 

print(replace_and_remove_doc, target = "replace_and_remove_doc.docx")

# edit replace_and_remove_doc
replace_and_remove_doc_edited <- read_docx(path = "replace_and_remove_doc.docx")  %>% 
        # remove paragraph with "drop that text"
        cursor_reach(keyword = "that text") %>% 
        body_remove() %>%

        # replace "test para 2" using pos = "on" argument of body_add_par
        cursor_reach(keyword = "that text") %>% 
        body_add_par(value = rep("test para 3", times = 3), style = "centered", pos = "on")

print(my_doc, target = "assets/docx/ipsum_doc.docx")



# you can use bookmarks, but they don't seem terribly useful
# the bookmark is assigned to the first "chunk" in the element where the cursor is
# then you can replace/remove this chunk using this bookmark
# even though you can have one paragraph element, it could still be composed to two or more chunks
# the number of chunks is based on how the paragrpah element was constructed (eg slip_in_text creates 2nd chunk)
# note that in the example below, the cursor defaults to the beginning of the element,
# and so the bookmark is placed on the first chunk of the element
doc <- read_docx() %>%
        body_add_par("centered text", style = "centered") %>%
        # slip_in_text(". How are you", style = "strong") %>%
        slip_in_text(". How are you", style = "strong", pos = "before") %>%
        # docx_show_chunk()
        body_bookmark("text_to_replace") %>%
        # docx_show_chunk()
        body_replace_text_at_bkm("text_to_replace", "not left aligned") %>% 
        docx_show_chunk()
        print(target = "bookmark_example.docx")


# replace all function
read_docx() %>%
        body_add_par("Placeholder one") %>%
        body_add_par("Placeholder two") %>%
        slip_in_text(". created second chunk in lower case placeholder two paragraph") %>%
        slip_in_text(". created second chunk in upper case Placeholder Two paragraph") %>%
        # docx_show_chunk()
        
        # replace with only_at_cursor arg = TRUE
        # body_replace_all_text(old_value = "Placeholder", new_value = "new", only_at_cursor = TRUE) %>%
        
        # replace with only_at_cursor_arg = FALSE
        # body_replace_all_text(old_value = "Placeholder", new_value = "new", only_at_cursor = FALSE) %>%
        
        # replace ignoring case 
        # body_replace_all_text(old_value = "Placeholder", new_value = "new", only_at_cursor = FALSE,
        #                       ignore.case = TRUE) %>%
        
        # replace using regex
        body_replace_all_text(old_value = "[P|p|z]lace.*", new_value = "new_regex", only_at_cursor = FALSE) %>%

        docx_show_chunk() %>% 
        docx_summary() %>%
        print(target = "replace_all_example.docx")

        






