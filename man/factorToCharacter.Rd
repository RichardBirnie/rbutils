% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/rbutils.R
\name{factorToCharacter}
\alias{factorToCharacter}
\title{Take a data frame and convert all factor variables to character variables}
\usage{
factorToCharacter(df)
}
\arguments{
\item{df}{A data frame with at least one character column}
}
\value{
The same data frame but with factors converted to character
}
\description{
Take a data frame and convert all factor variables to character variables
}
\examples{
data(iris)
class(iris$Species)
[1] "factor"
iris = factorToCharacter(iris)
class(iris$Species)
[1] "character"
}

