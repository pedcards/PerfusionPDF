/*	PerfusionPDF
 *	Import perfusion record PDF reports
 */

#Requires AutoHotkey v2

fname := ".\data\Record 6_23_2022.pdf"
fname := ".\data\20250814192513-out.txt"
y := record(fname)
n := 0
demog := y.block("\R*Patient Data\R",1,0,"\R+Surgery Team\R",1)
team := y.block("\R+Surgery Team",1,1,"\R+Disposables\R",1)
onlinedata := y.block("\R+Online\s+Data\R",1,0,"\R+Cardioplegia\s+Values\R",1)
y.getOnlineData(onlinedata)

/*	====================================================================================
 */
Class record
{
	__New(fileIn) {
		if FileExist(fileIn) {
			SplitPath(fileIn,&fName,&path,&ext,&fNam)
			this.file := []
			this.file.path := path
			this.file.fName := fName
			this.file.fNam := fNam
			this.file.ext := ext
			this.file.fOut := A_Now ".txt"
			this.exe := ".\includes\pdftotext.exe"
			this.opts := "-table -eol unix -nopgbrk -margint 108 -marginb 90"

			this.readFile(fileIn)

			return this
		} else {
			return Error
		}
	}

	readFile(fileIn) {
		if (this.file.ext="txt") {
			this.text := FileRead(fileIn)
			return
		}
		RunWait(this.exe " " this.opts " `"" this.file.fName "`" " this.file.fOut , this.file.path,"Hide")
		txtIn := FileRead(this.file.path "\" this.file.fOut)
		txtIn := StrReplace(txtIn,"`n`n","`n")
		FileAppend(txtIn,this.file.path "\" A_Now "-out.txt")
		this.text := txtIn
		; FileDelete(this.file.fOut)
	}

	/*	Get a block using stRegX
	 *		BS = begin string
	 *		BO = position from start of haystack
	 *		BT = trim BS, true or false
	 *		ES = end string
	 *		ET = trim ES, true or false
	 *		N  = var for next offset
	 */
	block(BS:="",BO:=1,BT:=0,ES:="",ET:=0,&N:=0) {
		res := stRegX(this.text,BS,BO,BT,ES,ET,&N)
		return res
	}

	getOnlineData(txt) {
	/*	Parse txt block for Online Data fields
	 *	After header, scan for column labels and cell coords
	 *	Read each cell
	 */
		/*	First put all lines into an array
		*/
		txtarray := []
		loop parse txt, "`n" {
			txtarray.Push(A_LoopField)
		}

		/*	Read each line in sequence
		 *	"Online Data" triggers column rescan
		 *	"Time" should always be first column
		 *	Look ahead for time-contaning line,
		 *	may need to concat time + AM/PM lines
		 *	Store as timearray, objects indexed to timestamp	
		 */
		timearray := []
		for idx in txtarray
		{
			txtline := txtarray[A_Index]
			if (txtline~="Online\s+Data") {												; Line is Online Data
				cols := getHeaders(A_Index)
				A_Index := A_Index + cols.height
				continue
			}
		}

		getHeaders(idx) {
		/*	Get column names and X coords
		*/
			hdrIdx := idx+1																; Header is first line after "Online Data"
			line := txtarray[hdrIdx]
			line := RegExReplace(line,"  P  "," P   ")									; Adjust P in headers
			colx := findHeaders(line)													; Get positions of first row of header
			
			header := Map()
			loop 6 {																	; Look ahead next 6 lines
				row := A_Index
				x := readRow(txtarray[idx+row],colx)
				if (x[1]~="^\d{1,2}:\d{2}") {											; Until finds time string in first cell
					for key in header {
						tx := header[key]
						cleanspace(&tx)
						tx := RegExReplace(tx,"(\[.*?\])")
						header[key] := Trim(tx)
					}
					return {hdr:header,col:colx,height:row-1}
				}
				for key,val in x {														; Concatenate lines of header
					tx := x[key]
					if (row=1) {
						header[key] := tx
					} else {
						header[key] .= tx
					}
				}
			}
		}

		findHeaders(txt) {
		/*	Match position of  first char after "\s\s"
		*/
			n := 0
			res := []
			loop {
				n := RegExMatch(txt,"^\S|(?<=\s\s)\S",&match,n+1)
				if !n {
					break
				}
				res.Push(match.pos)
			}
			return res
		}

		readRow(txt,colx) {
		/*	Return array of chunks from colx
		*/
			res := Map()
			for key,val in colx {
				x1 := colx[key]
				ln := (key<colx.length) 
					? colx[key+1]-x1 
					: StrLen(txt)-x1+1
				tx := SubStr(txt,x1,ln)
				cleanspace(&tx)
				res[key] := trim(tx)
			}
			return res
		}
	}
}

cleanspace(&txt) {
	txt := StrReplace(txt,"`n"," ")
	txt := StrReplace(txt," . ",". ")
	loop 
	{
		txt := StrReplace(txt,"  "," ",, &count)
		if (count=0)	
			break
	}
}

#Include includes
#Include strx2.ahk
#Include Peep.v2.ahk