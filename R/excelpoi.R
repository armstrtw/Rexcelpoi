write.xls <- function(x,
                      filename,
                      sheetName="fromR",
                      formats=NULL)
{
    if(typeof(x)=="list" && !is.data.frame(x)) {
        sheetName <- NULL
    }

    invisible(.Call("R2xls",x,filename,sheetName,formats,PACKAGE="Rexcelpoi"))
}
