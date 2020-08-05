library(devtools)
install.packages("goodpractice")
usethis::use_git()
use_r(name = "Yes_MCA")
load_all()
document()

#use_build_ignore("dev_history.R")

#use_gpl3_license("Sébastien Lê")

use_tidy_description()

attachment::att_to_description()

check()

goodpractice::goodpractice()

install()

use_github()

#  57748b7b4441187694e3e97b4cf8f20e39a5ebb8


