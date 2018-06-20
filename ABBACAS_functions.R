# store custom functions
# author: ABBA Consulting and Services
# require this at start of forecasting scripts
# keep master copy on OneDrive - copy instance to project folders
# goldmaster: C:\Users\SQLmaster\OneDrive\Documents\ABBACAS\R_functions\ABBACAS_functions.R


# first a create a function to read in an Excel file and store back as CSV file
read_then_csv <- function(sheet, path, columntypes, columnnames) {
  pathbase <- path %>%
    basename() %>%
    tools::file_path_sans_ext()
  path %>%
    read_excel(sheet = sheet, 
               na = c("",".","NA","N/A","#N/A"," ","#VALUE!"),
               col_types = columntypes,
               col_names = poscolumnnames) %>%
    write_csv(paste(ddir,paste0(pathbase, "-", sheet, ".csv"),sep="/"))
}

#generate date-array matched to min(week) to max(week) for time-series:
# min_week <- min(week_variable)
# max_week <- max(week_variable)
# n_weeks <- (max_week - min_week)/7
setdatesarray <- function(min_week = "2012-01-01", max_week ="2017-12-25") {
  min_week = ymd(min_week)
  max_week = ymd(max_week)
  n_weeks =ceiling(as.numeric(max_week - min_week)/7.0 + 1)
  return(seq(as.Date(min_week), by = "weeks", length.out = n_weeks))
}

setdatesarray2 <- function(timearray) {
  min_week = ymd(min(timearray))
  max_week = ymd(max(timearray))
  n_weeks =ceiling(as.numeric(max_week - min_week)/7.0 + 1)
  return(seq(as.Date(min_week), by = "weeks", length.out = n_weeks))
}


# Writing Output to Reporting Directory -----------------------------------

# makes a subdirectory relative to the current (project) directory
makeprojectsubdir <- function(subdir) {
  chkwd <- getwd() # this should be the project directory and where r-scripts are stored
  subdir  <- paste(chkwd, subdir, sep = "/")
  if(!dir.exists(subdir)) dir.create(subdir) #creates data diretory if not already existing
  return((subdir))
}

# paste directory to file name
fullpathname <- function(rdir, fname){
  as.character(paste(rdir, fname, sep = "/"))
}

# output graph to PDF file
savePDF <- function(myPlot, rdir, namePlot) {
  fname = (paste0(namePlot, ".pfd"))
  fpname = fullpathname(rdir, fname)
  pdf(fpname, useDingbats = FALSE)
  print(myPlot)
  dev.off()
}

# special case: output "checkresiduals()" graph to jpg file
save_residualsJPG <- function(fcst, rdir, fname) {
  fname = paste0(fname, ".jpg")
  fpname = fullpathname(rdir, fname)
  jpeg(file = fpname, width = 1080, height = 720)
  print(checkresiduals(fcst))
  dev.off()
}

# output graph to jpg file
saveJPG <- function(myPlot, rdir, fname) {
  fname = (paste0(fname, ".jpg"))
  fpname = fullpathname(rdir, fname)
  jpeg(file = fpname)
  print(myPlot)
  dev.off()
}

# output graph to png file
savePNG <- function(myPlot, rdir, fname) {
  fname = (paste0(nameJPG, ".png"))
  fpname = fullpathname(rdir, fname)
  png(file = fpname)
  print(myPlot)
  dev.off()
}

# output graph to TIFF file
saveTIFF <- function(myPlot, rdir, fname) {
  fname = (paste0(fname, ".tif"))
  fpname = fullpathname(rdir, fname)
  tiff(file = fpname)
  print(myPlot)
  dev.off()
}

# output graph to BMP file
saveBMP <- function(myPlot, rdir, fname) {
  fname = (paste0(fname, ".jpg"))
  fpname = fullpathname(rdir, fname)
  bmp(file = fpname)
  print(myPlot)
  dev.off()
}


# This should enable output of forecasting graphics into a PPTX deck
# This function creates 2-up image slide and can be called from loops

TwoContent <- function(rdir, dname, dtitle, dsub1, dsub2){
  fname1 = paste0(dsub1,dname)
  fname2 = paste0(dsub2,dname)
  doc <- addSlide(doc, "Two Content")
  doc <- addTitle(doc, paste0(dtitle, dname))
  doc <- addImage(doc, paste0(fullpathname(rdir, fname1),".jpg"))
  doc <- addImage(doc, paste0(fullpathname(rdir, fname2),".jpg"))
  doc <- addDate(doc)
  doc <- addFooter(doc, "Fridababy")
  doc <- addPageNumber(doc)
}


# this is a single graphic with Title
# Silde 4 : Add R script
#+++++++++++++++++++++
OneContent <- function(rdir, dname, dtitle, dsub3){
  fname3 = paste0(dsub3,dname)
doc <- addSlide(doc, "Title and Content")
doc <- addTitle(doc, paste0(dtitle, dname))
doc <- addImage(doc, paste0(fullpathname(rdir, fname3),".jpg"))
doc <- addDate(doc)
doc <- addFooter(doc, "Fridababy")
doc <- addPageNumber(doc)
}