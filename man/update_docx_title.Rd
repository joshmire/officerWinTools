\name{update_docx_title}
\alias{update_docx_title}
\title{Update the title property of a .docx file created with the 'officer' package}
\description{
This function exports the document as a temporary .docx file, creates a
temporary .vbs file to update the document's title property and save, runs
the .vbs file in command prompt, overrides x using the 'read_docx()' function
of the 'officer' package, deletes the temporary files, and returns the
updated .docx file.
}
\usage{
update_docx_title(x, title = NULL)
}
\arguments{
\item{x}{document created using 'read_docx()' function from 'officer' package}
\item{title}{string the document's title property is to be set to}
}
\examples{
doc <- read_docx()
doc <- update_docx_title(doc, "Example title")
}
