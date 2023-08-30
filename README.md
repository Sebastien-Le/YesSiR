
<!-- README.md is generated from README.Rmd. Please edit that file -->

# YesSiR

<!-- badges: start -->
<!-- badges: end -->

The *YesSiR* package is a follow-up to the *EnQuireR* package. You can
find some information about the *EnQuireR* package by clicking on this
[link](http://enquirer.free.fr/index.html) or on this
[one](https://www.r-project.org/conferences/useR-2009/booklet.pdf); if
you want to download it, go to the [CRAN](https://cran.r-project.org)
(Packages, then Archive). The main objective of the *EnQuireR* package
was the creation of automatic reports for survey data. The automatic
generation of reports was carried out using Sweave and Latex
([link1](https://stat.ethz.ch/R-manual/R-devel/library/utils/doc/Sweave.pdf),
[link2](https://research.wu.ac.at/en/publications/sweave-dynamic-generation-of-statistical-reports-using-literate-d-3),
[link3](https://rpubs.com/YaRrr/SweaveIntro),
[link4](https://support.posit.co/hc/en-us/articles/200552056-Using-Sweave-and-knitr)):
as you can see, thanks to the work of Friedrich Leisch, it’s been two
decades that statistical reports can be generated with R.

Since then, technology has evolved and it’s hard to imagine a
presentation without PowerPoint. The aim of the *YesSiR* package is to
generate automatic reports in PowerPoint format. It works thanks to the
work of David Gohel and his famous package
[*officer*](https://davidgohel.github.io/officer/). The functions in the
*YesSiR* package take as input the results of the `PCA()`, `MCA()` and
`textual()` functions in the *FactoMineR* package, as well as the
results of the `decat()` function in the *SensoMineR* package. These
functions produce a report in the form of a PowerPoint presentation
which is automatically saved in the working directory (by default).

This package would not have been possible without the hard work and
efficiency of Maxime Saland. The main purpose of the package is to
demonstrate what can be achieved in terms of automatic reporting.
Adjustments will need to be made to suit specific practices and
applications.

## Installation

You can install the released version of YesSiR from
[GitHub](https://github.com/) with:

``` r
# install.packages("devtools")
devtools::install_github("Sebastien-Le/YesSiR")
```

## Output

``` r
# Generation of the results you want to put in your PowerPoint
library(FactoMineR)
data(decathlon)
res.pca <- FactoMineR::PCA(decathlon, quanti.sup = 11:12, quali.sup=13)

# Creation of the PowerPoint in the working directory
library(YesSiR)
Yes_PCA(res.pca)
```

If you want to have a look at the PowerPoint presentation that has been
created, you can download it
[here](https://github.com/Sebastien-Le/YesSiR/blob/master/PCA_results.pptx)
(small icon on your right).
