library(devtools)
usethis::use_git()
use_r(name = "Yes_MCA")
load_all()
document()

#use_build_ignore("dev_history.R")

#use_gpl3_license("Sébastien Lê")

use_tidy_description()

attachment::att_to_description()

check()
