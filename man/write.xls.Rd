\name{write.xls}
\alias{write.xls}
\title{Write MS excel files}
\description{
 Create new excel files from data in R.
}
\usage{
write.xls(x,filename,sheetName,formats=NULL)
}
\arguments{
	\item{x}{a dataframe, matrix, vector, or list containing those objects}
	\item{filename}{the filename of the output file}
	\item{sheetName}{the name of the excel sheet that will be created}
	\item{formats}{a vector of strings used to format each column of
          the excel sheet}
}

\value{T or F indicatting success or failure}

\examples{
nc <- 4
nr <- 20
x <- matrix(rnorm(nc*nr),ncol=nc,nrow=nr)

rownames(x) <- rep(letters,length=nr)
colnames(x) <- rep(letters,length=nc)

# write a matrix
write.xls(x,"matrixFromR.xls","matrix")

# add some formats
write.xls(x,"matrixFromR_formatted.xls","matrix",formats=rep("0.0",nc))
write.xls(x,"matrixFromR_formatted2.xls","matrix",formats=rep(c("0.0","0"),nc/2))
write.xls(x,"matrixFromR_formatted3.xls","matrix",formats=rep("#,##0.0;[Red]#-##0.0;\"\"",nc))

write.xls(x[,1],"vectorFromR.xls","vector")

# char data test
y <- matrix(rep(letters,each=26),ncol=26)
rownames(y) <- rep(letters,length=26)
colnames(y) <- rep(letters,length=26)
write.xls(y,"char.data.xls","alphabet")


# logical data test
z <- matrix(as.logical(round(runif(nc*nr),0)),ncol=nc)
rownames(z) <- rep(letters,length=nr)
colnames(z) <- rep(letters,length=nc)

write.xls(z,"logical.data.xls","myLogic")
write.xls(z[,1],"logical.data.vector.xls","myLogicVector")

# test for writing data frames
x.df <- as.data.frame(x)
write.xls(x.df,"data.frame.xls","myDF")

# test list writing

data.list <- list()
data.list[["double"]] <- x
data.list[["char"]] <- y
data.list[["logical"]] <- z
data.list[["dataframe"]] <- x.df

## accepts list of formats
fmt.list <- list("0.00","","0","0.0000")

write.xls(data.list,"listFromR.xls",formats=fmt.list)

## nested list for multiple objects on the same excel sheet
nested <- list()
nested[["same.page"]] <- data.list
nested[["normal"]] <- x
write.xls(nested,"same.page.xls",formats=list(sampe.page=list("0.00","","0","0.0"),normal="0.000"))
}

\references{
 http://jakarta.apache.org/poi/index.html
}
\details{
	Not completed yet
}

\author{Whit Armstrong \email{armstrong.whit@gmail.com}}
\keyword{IO}
