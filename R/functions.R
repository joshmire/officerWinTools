packageStartupMessage('Warning:  The "officerWinTools" package relies upon a Windows operating system in order to run VBScript in Command Prompt.')

#' Export a .docx file created using 'officer' package as a .pdf
#'
#' This function exports the document as a temporary .docx file, creates a
#' temporary .vbs file with the intended .pdf file name, runs the .vbs file in
#' command prompt, and then deletes the temporary files.
#'
#' @param x document created using 'read_docx()' function from 'officer' package
#' @param target path to the .pdf to be exported, default is a .pdf with the exported object's name in the current working directory
#' @examples
#' doc1 <- read_docx()
#' print_docx_pdf(doc1)
#' doc2 <- read_docx()
#' doc2 <- body_add_par(doc2, value = "Table of content", style = "heading 1")
#' doc2 <- body_add_toc(doc2, level = 2)
#' doc2 <- body_end_section_continuous(doc2)
#' doc2 <- body_add_par(doc2, value = "Section 1", style = "heading 1")
#' doc2 <- body_add_par(doc2, value = "This is a test.", style = "heading 2")
#' doc2 <- body_add_par(doc2, value = "Section 1", style = "heading 1")
#' print_docx_pdf(doc2, target = file.path(paste(getwd(),"/test.pdf",sep="")))
#' @export
print_docx_pdf <- function(x, target = NULL, ...){

	if(Sys.info()['sysname'] != "Windows"){
		stop('The "officerWinTools" package requires Windows operating system.  "print_docx_pdf" will only work in a Windows enviroment.')
	}

  if(is.null(target)){
    target <- file.path(paste(getwd(),"/",deparse(substitute(x)),".pdf",sep=""))
  }

	if( !grepl(x = target, pattern = "\\.(pdf)$", ignore.case = TRUE) ){
		stop(target , " should have '.pdf' extension.")
	}

	invisible(suppressWarnings(R.utils::mkdirs(normalizePath(dirname(target)))))

	print(x, target = file.path(paste(getwd(),"/temp.docx",sep="")))

	writeLines(
		c(
			'Set objWord = CreateObject("Word.Application")',
			'objWord.Visible = False',
			'objWord.DisplayAlerts = False',
			paste('Set doc = objWord.Documents.Open("',normalizePath(paste(getwd(),"/temp.docx",sep="")),'")',sep=""),
			'On Error Resume Next',
			'For Each TOC In doc.TablesOfContents',
			'TOC.Update',
			'Next',
			'On Error GoTo 0',
			paste('Call doc.SaveAs2("',target,'", 17)',sep=""),
			'doc.Saved = TRUE',
			'objWord.Quit'

		),
		con = file.path(paste(getwd(),"/temp.vbs",sep="")),
		sep = "\n",
		useBytes = FALSE
	)

	shell(shQuote(normalizePath(file.path(paste(getwd(),"/temp.vbs",sep="")))), "cscript", flag = "//nologo")

	invisible(file.remove(file.path(paste(getwd(),"/temp.docx",sep=""))))

	invisible(file.remove(file.path(paste(getwd(),"/temp.vbs",sep=""))))
}


#' Export a .pptx file created using 'officer' package as a .pdf
#'
#' This function exports the document as a temporary .pptx file, creates a
#' temporary .vbs file with the intended .pdf file name, runs the .vbs file in
#' command prompt, and then deletes the temporary files.
#'
#' @param x document created using 'read_pptx()' function from 'officer' package
#' @param target path to the .pdf to be exported, default is a .pdf with the exported object's name in the current working directory
#' @examples
#' ppt1 <- read_pptx()
#' ppt1 <- add_sheet(ppt1)
#' print_pptx_pdf(ppt1)
#' ppt2 <- read_pptx()
#' ppt2 <- add_sheet(ppt2)
#' print_pptx_pdf(ppt2, target = file.path(paste(getwd(),"/test.pdf",sep="")))
#' @export
print_pptx_pdf <- function(x, target = NULL, ...){

	if(Sys.info()['sysname'] != "Windows"){
		stop('The "officerWinTools" package requires Windows operating system.  "print_pptx_pdf" will only work in a Windows enviroment.')
	}

  if(is.null(target)){
    target <- file.path(paste(getwd(),"/",deparse(substitute(x)),".pdf",sep=""))
  }

	if( !grepl(x = target, pattern = "\\.(pdf)$", ignore.case = TRUE) ){
		stop(target , " should have '.pdf' extension.")
	}

	invisible(suppressWarnings(R.utils::mkdirs(normalizePath(dirname(target)))))

	print(x, target = file.path(paste(getwd(),"/temp.pptx",sep="")))

	writeLines(
		c(
			'Set objPpt = CreateObject("Powerpoint.Application")',
			'objPpt.DisplayAlerts = False',
			paste('Set pres = objPpt.Presentations.Open("',normalizePath(paste(getwd(),"/temp.pptx",sep="")),'")',sep=""),
			'On Error Resume Next',
			paste('Call pres.SaveAs("',target,'", 32)',sep=""),
			'On Error GoTo 0',
			'pres.Saved = TRUE',
			'objPpt.Quit'

		),
		con = file.path(paste(getwd(),"/temp.vbs",sep="")),
		sep = "\n",
		useBytes = FALSE
	)

	shell(shQuote(normalizePath(file.path(paste(getwd(),"/temp.vbs",sep="")))), "cscript", flag = "//nologo")

	invisible(file.remove(file.path(paste(getwd(),"/temp.pptx",sep=""))))

	invisible(file.remove(file.path(paste(getwd(),"/temp.vbs",sep=""))))
}


#' Update the tables of contents of a .docx file created with the 'officer' package
#'
#' This function exports the document as a temporary .docx file, creates a
#' temporary .vbs file to update all references of all tables of contents and
#' save, runs the .vbs file in command prompt, overrides x using the
#' 'read_docx()' function of the 'officer' package, deletes the temporary files,
#'  and returns the updated .docx file.
#'
#' @param x document created using 'read_docx()' function from 'officer' package
#' @examples
#' doc <- read_docx()
#' doc <- body_add_par(doc, value = "Table of content", style = "heading 1")
#' doc <- body_add_toc(doc, level = 2)
#' doc <- body_end_section_continuous(doc)
#' doc <- body_add_par(doc, value = "Section 1", style = "heading 1")
#' doc <- body_add_par(doc, value = "This is a test.", style = "heading 2")
#' doc <- body_add_par(doc, value = "Section 1", style = "heading 1")
#' doc <- update_toc(doc)
#' @export
update_toc <- function(x){

  if(Sys.info()['sysname'] != "Windows"){
    stop('The "officerWinTools" package requires Windows operating system.  "update_toc" will only work in a Windows enviroment.')
  }

  tryCatch(
    {print(x, target = file.path(paste(getwd(),"/temp.docx",sep="")))},
    error=function(cond) {stop("x is not a Word document.")}
  )

  print(x, target = file.path(paste(getwd(),"/temp.docx",sep="")))

  writeLines(
    c(
      'Set objWord = CreateObject("Word.Application")',
      'objWord.Visible = False',
      'objWord.DisplayAlerts = False',
      paste('Set doc = objWord.Documents.Open("',normalizePath(paste(getwd(),"/temp.docx",sep="")),'")',sep=""),
      'On Error Resume Next',
      'For Each TOC In doc.TablesOfContents',
      'TOC.Update',
      'Next',
      'On Error GoTo 0',
      'doc.Close (TRUE)',
      'objWord.Quit'

    ),
    con = file.path(paste(getwd(),"/temp.vbs",sep="")),
    sep = "\n",
    useBytes = FALSE
  )

  shell(shQuote(normalizePath(file.path(paste(getwd(),"/temp.vbs",sep="")))), "cscript", flag = "//nologo")

  x <- officer::read_docx(path = file.path(paste(getwd(),"/temp.docx",sep="")))

  invisible(file.remove(file.path(paste(getwd(),"/temp.docx",sep=""))))

  invisible(file.remove(file.path(paste(getwd(),"/temp.vbs",sep=""))))

  x
}
