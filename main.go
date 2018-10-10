package main

import (
	"fmt"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

type energia struct {
	przec           string
	zaklad          int
	po              string
	data            time.Time //data faktury w raporcie
	przypisanie     string
	nrDok           string
	rodzaj          string
	dataDokumentu   time.Time
	referencja      string
	kk              int
	kwotaWkr        float64
	walKr           string
	ilosc           float64
	pd              string
	opis            string
	fakOkresowa     bool
	dateFrom        time.Time //data faktury okresowej OD
	dateTo          time.Time //data faktury okresowej DO
	accountingMonth int       //miesiac do ktorego nalezy zaliczyc fakture ksiegowo
	accountingYear  int       //rok do ktorego nalezy zaliczyc fakture ksiegowo
}

var (
	wiersz    energia
	endOfRows int
	fileObr   *xlsx.File
	fileSpz   *xlsx.File
	plantNbr  int
	sheet     *xlsx.Sheet
	row       *xlsx.Row
	cell      *xlsx.Cell
	col       *xlsx.Col
	v         int
)

func checkInvoiceTimePeriod(wiersze []energia) []energia {
	for i := range wiersze {
		monthArr := strings.Split((strings.Split(wiersze[i].opis, ",")[1]), "/")
		if len(monthArr) < 3 {
			wiersze[i].fakOkresowa = false
		} else {
			wiersze[i].fakOkresowa = true
		}
	}
	return wiersze
}

func setDateFromDateTo(wiersze []energia) []energia {
	for i := range wiersze {
		if wiersze[i].fakOkresowa == true { //faktura okresowa
			monthArr := strings.Split((strings.Split(wiersze[i].opis, ",")[1]), "/")
			if len(monthArr) > 3 {
				dayFrom, _ := strconv.Atoi(monthArr[0])
				monthFrom, _ := strconv.Atoi(monthArr[1])
				yearFrom, _ := strconv.Atoi(strings.Split(monthArr[2], "-")[0])
				dayTo, _ := strconv.Atoi(strings.Split(monthArr[2], "-")[1])
				monthTo, _ := strconv.Atoi(monthArr[3])
				yearTo, _ := strconv.Atoi(monthArr[4])

				wiersze[i].dateFrom = time.Date(yearFrom, time.Month(monthFrom), dayFrom, 0, 0, 0, 0, time.UTC)
				wiersze[i].dateTo = time.Date(yearTo, time.Month(monthTo), dayTo, 0, 0, 0, 0, time.UTC)
			}
		}
	}
	return wiersze
}

func setAccountingMonth(wiersze []energia) []energia {
	for i := range wiersze {
		if wiersze[i].referencja == "97971/1804/00172" {
			fmt.Println(wiersze[i].referencja)
		}

		if wiersze[i].fakOkresowa == false { //faktura miesieczna
			monthArr := strings.Split((strings.Split(wiersze[i].opis, ",")[1]), "/")
			if len(monthArr) > 2 {
				wiersze[i].accountingMonth, _ = strconv.Atoi(monthArr[1])
				wiersze[i].accountingYear, _ = strconv.Atoi(monthArr[2])
			} else {
				wiersze[i].accountingMonth, _ = strconv.Atoi(monthArr[0])
				wiersze[i].accountingYear, _ = strconv.Atoi(monthArr[1])
			}

		} else if wiersze[i].fakOkresowa == true { //faktura okresowa
			day := wiersze[i].dateTo.Day()
			if int(day) == 1 {
				wiersze[i].accountingMonth = int(wiersze[i].dateTo.Month()) - 1 //miesiac wczesniej okres podany w formacie do pierwszego dnia nowego miesiaca
			} else {
				wiersze[i].accountingMonth = int(wiersze[i].dateTo.Month())
			}
			wiersze[i].accountingYear = int(wiersze[i].dateTo.Year())

		}
	}

	return wiersze
}

func splitSlicePerPlant(wiersze []energia) ([][]energia, int) {
	wierszePerPlant := make([][]energia, 0)

	plantNbr = wiersze[0].zaklad
	v, m, z, i := 0, 0, 0, 0
	for i = range wiersze {
		if plantNbr != wiersze[i].zaklad {
			fmt.Println(wiersze[i].zaklad)
			z += m
			l := i
			k := z - m
			plantNbr = wiersze[i].zaklad
			wierszePerPlant = append(wierszePerPlant, wiersze[k:l])
			v++
			m = 1
		} else {
			m++
		}
	}
	//dodatnie ostatniej pozycji - ostani zaklad
	z += m
	l := i + 1
	k := z - m
	plantNbr = wiersze[i].zaklad
	wierszePerPlant = append(wierszePerPlant, wiersze[k:l])

	return wierszePerPlant, v
}

func invoiceAggregation(wiersze []energia) []energia {

	wierszeNew := make([]energia, 0)
	var nrFaktury string
	var kwotaTemp float64
	var iloscTemp float64
	var exist bool

	for i := range wiersze {
		exist = false
		nrFaktury = wiersze[i].referencja
		kwotaTemp = 0
		iloscTemp = 0
		for j := range wiersze {
			if nrFaktury == wiersze[j].referencja {
				kwotaTemp += wiersze[j].kwotaWkr
				iloscTemp += wiersze[j].ilosc
			}
		}
		for k := range wierszeNew {
			if nrFaktury == wierszeNew[k].referencja {
				exist = true
			}
		}
		if exist == false {
			wierszeTemp := wiersze[i]
			wierszeTemp.kwotaWkr = kwotaTemp
			wierszeTemp.ilosc = iloscTemp
			wierszeNew = append(wierszeNew, wierszeTemp)

		}
	}

	return wierszeNew
}

func saveToNewSheet(wiersze []energia, typ string, fileObr xlsx.File) xlsx.File {

	//wiersze = invoiceAggregation(wiersze)

	const shortForm = "2006-Jan-02"

	xlsx.SetDefaultFont(9, "Calibri")
	headStyle := xlsx.NewStyle()
	headStyle.Alignment.WrapText = true
	headStyle.Alignment.Horizontal = "center"
	headStyle.Alignment.Vertical = "center"
	headStyle.Font.Size = 9
	headStyle.Font.Color = "00FFFFFF"
	headStyle.Font.Bold = true
	headStyle.Border = *xlsx.NewBorder("medium", "medium", "medium", "medium")
	headStyle.Fill = *xlsx.NewFill("solid", "00B30000", "00B30000")

	col0Style := xlsx.NewStyle()
	col0Style.Alignment.WrapText = true
	col0Style.Alignment.Horizontal = "center"
	col0Style.Alignment.Vertical = "center"
	col0Style.Font.Bold = true
	col0Style.Border = *xlsx.NewBorder("thin", "thin", "thin", "thin")

	col0StyleLeft := xlsx.NewStyle()
	col0StyleLeft.Alignment.WrapText = true
	col0StyleLeft.Alignment.Horizontal = "center"
	col0StyleLeft.Alignment.Vertical = "center"
	col0StyleLeft.Font.Bold = true
	col0StyleLeft.Border = *xlsx.NewBorder("medium", "thin", "thin", "thin")

	col0StyleRight := xlsx.NewStyle()
	col0StyleRight.Alignment.WrapText = true
	col0StyleRight.Alignment.Horizontal = "center"
	col0StyleRight.Alignment.Vertical = "center"
	col0StyleRight.Font.Bold = true
	col0StyleRight.Border = *xlsx.NewBorder("thin", "medium", "thin", "thin")

	// col10Style := xlsx.NewStyle()
	// col10Style.Alignment.WrapText = true
	// col10Style.Alignment.Horizontal = "center"
	// col10Style.Alignment.Vertical = "center"
	// col10Style.Font.Bold = true
	// col10Style.Border = *xlsx.NewBorder("thin", "thin", "thin", "thin")

	headStyleNW := xlsx.NewStyle()
	headStyleNW.Alignment.WrapText = false
	headStyleNW.Font.Bold = true

	sumStyle := xlsx.NewStyle()
	sumStyle.Alignment.Horizontal = "center"
	sumStyle.Alignment.Vertical = "center"
	sumStyle.Font.Bold = true
	sumStyle.Fill = *xlsx.NewFill("solid", "00FFBBBB", "00FFBBBB")
	sumStyle.Border = *xlsx.NewBorder("thin", "thin", "thin", "thin")

	sumStyleLeft := xlsx.NewStyle()
	sumStyleLeft.Alignment.Horizontal = "center"
	sumStyleLeft.Alignment.Vertical = "center"
	sumStyleLeft.Font.Bold = true
	sumStyleLeft.Fill = *xlsx.NewFill("solid", "00FFBBBB", "00FFBBBB")
	sumStyleLeft.Border = *xlsx.NewBorder("medium", "thin", "thin", "thin")

	sumStyleRight := xlsx.NewStyle()
	sumStyleRight.Alignment.Horizontal = "center"
	sumStyleRight.Alignment.Vertical = "center"
	sumStyleRight.Font.Bold = true
	sumStyleRight.Fill = *xlsx.NewFill("solid", "00FFBBBB", "00FFBBBB")
	sumStyleRight.Border = *xlsx.NewBorder("thin", "medium", "thin", "thin")

	defCellStyle := xlsx.NewStyle()
	defCellStyle.Border = *xlsx.NewBorder("thin", "thin", "thin", "thin")
	defCellStyle.Alignment.Vertical = "center"
	defCellStyle.Alignment.Horizontal = "center"

	warning := xlsx.NewStyle()
	warning.Border = *xlsx.NewBorder("thin", "thin", "thin", "thin")
	warning.Alignment.Vertical = "center"
	warning.Alignment.Horizontal = "center"
	warning.Fill = *xlsx.NewFill("solid", "00FFFF00", "00FFFF00")

	sumTotalStyle := xlsx.NewStyle()
	sumTotalStyle.Alignment.Horizontal = "center"
	sumTotalStyle.Alignment.Vertical = "center"
	sumTotalStyle.Font.Bold = true
	sumTotalStyle.Fill = *xlsx.NewFill("solid", "00D7D8D9", "00D7D8D9")
	sumTotalStyle.Border = *xlsx.NewBorder("thin", "thin", "thin", "medium")

	sumTotalStyleLeft := xlsx.NewStyle()
	sumTotalStyleLeft.Alignment.Horizontal = "center"
	sumTotalStyleLeft.Alignment.Vertical = "center"
	sumTotalStyleLeft.Font.Bold = true
	sumTotalStyleLeft.Fill = *xlsx.NewFill("solid", "00D7D8D9", "00D7D8D9")
	sumTotalStyleLeft.Border = *xlsx.NewBorder("medium", "thin", "thin", "medium")

	sumTotalStyleRight := xlsx.NewStyle()
	sumTotalStyleRight.Alignment.Horizontal = "center"
	sumTotalStyleRight.Alignment.Vertical = "center"
	sumTotalStyleRight.Font.Bold = true
	sumTotalStyleRight.Fill = *xlsx.NewFill("solid", "00D7D8D9", "00D7D8D9")
	sumTotalStyleRight.Border = *xlsx.NewBorder("thin", "medium", "thin", "medium")

	plantNbr = 0
	sheet, err := fileObr.AddSheet(strconv.Itoa(wiersze[0].zaklad))

	currnetMonth := ""
	currentPeriod := ""
	rowNbr := 0
	rowInMonth := 0
	totalkWhAll := 0.0
	totalPLNAll := 0.0
	v := 0
	for i := range wiersze {
		if plantNbr != wiersze[i].zaklad {
			plantNbr = wiersze[i].zaklad
			if err != nil {
				fmt.Printf(err.Error())
			}
			for k := 0; k < 10; k++ {
				col = sheet.Col(k)
				col.Width = 15
			}
			//custom col width
			col = sheet.Col(4)
			col.Width = 20
			col = sheet.Col(5)
			col.Width = 35
			col = sheet.Col(10)
			col.Width = 20
			row = sheet.AddRow()
			cell = row.AddCell()
			cell = row.AddCell()
			data := wiersze[i].data
			//data = strings.Split(data, ".")[2]
			//data = "January " + data + " - " + "December " + data
			year := data.Year()
			cell.Value = "January " + strconv.Itoa(year) + " - " + "December " + strconv.Itoa(year)
			cell.SetStyle(headStyleNW)
			cell = row.AddCell()
			cell = row.AddCell()
			cell = row.AddCell()
			cell.Value = strconv.Itoa(wiersze[i].zaklad)
			cell.SetStyle(headStyleNW)
			row = sheet.AddRow()
			row = sheet.AddRow()
			cell = row.AddCell()
			//sheet Header
			cell = row.AddCell()
			cell.Value = "okres którego dotyczy faktura"
			cell.SetStyle(headStyle)
			cell = row.AddCell()
			cell.Value = "nr dowodu"
			cell.SetStyle(headStyle)
			cell = row.AddCell()
			cell.Value = "data faktury"
			cell.SetStyle(headStyle)
			cell = row.AddCell()
			cell.Value = "nr faktury"
			cell.SetStyle(headStyle)
			cell = row.AddCell()
			cell.Value = ""
			cell.SetStyle(headStyle)
			cell = row.AddCell()
			cell.Value = "Total Electricity use [kWh]"
			cell.SetStyle(headStyle)
			cell = row.AddCell()
			cell.Value = "Total Cost, Energy [PLN - netto]"
			cell.SetStyle(headStyle)
			cell = row.AddCell()
			cell.Value = "Total Cost, Energy [PLN/kWh]"
			cell.SetStyle(headStyle)
			cell = row.AddCell()
			cell.Value = "dostawca"
			cell.SetStyle(headStyle)
			cell = row.AddCell()
			cell.Value = "nazwa dostawcy"
			cell.SetStyle(headStyle)
			rowNbr = 3
		}
		// Wyklucznie grudnia roku poprzedniego
		if plantNbr == wiersze[i].zaklad && wiersze[i].przypisanie == typ && wiersze[i].accountingMonth != 12 && wiersze[i].accountingYear != 2017 {
			v = i //to track last index for certain type
			rowInMonth++
			rowNbr++
			month := strings.Split((strings.Split(wiersze[i].opis, ",")[1]), "/")[0]
			//monthArr := strings.Split((strings.Split(wiersze[i].opis, ",")[1]), "/")
			// if len(monthArr) < 3 {
			// 	month = monthArr[0]
			// } else {
			// 	month = monthArr[3]
			// }
			month = strconv.Itoa(wiersze[i].accountingMonth)
			if currnetMonth != "" && currnetMonth != month { //podsumowanie miesiaca
				row = sheet.AddRow()
				cell = row.AddCell()
				cell = row.AddCell()
				cell.Value = currentPeriod
				cell.SetStyle(sumStyleLeft)
				currnetMonth = month
				currentPeriodArr := strings.Split(wiersze[i].opis, ",")
				if len(strings.Split(currentPeriodArr[1], "/")) < 3 {
					//currentPeriod = strings.Split(wiersze[i].opis, ",")[1]
					currentPeriod = strconv.Itoa(wiersze[i].accountingMonth) + "/" + strconv.Itoa(wiersze[i].accountingYear)
				} else {
					//currentPeriod = strings.Split(strings.Split(strings.Split(wiersze[i].opis, ",")[1], "-")[1], "/")[1] + "/" + strings.Split(strings.Split(strings.Split(wiersze[i].opis, ",")[1], "-")[1], "/")[2]
					currentPeriod = strconv.Itoa(wiersze[i].accountingMonth) + "/" + strconv.Itoa(wiersze[i].accountingYear)
				}
				cellCoordinate := xlsx.GetCellIDStringFromCoords(7, rowNbr-3) //wykonywane w wierszu ponizej podsumowania (inny miesiac) + wiersz podsumowujacy, wiersze i kolumny numerwoane od 0 w go dlatego -3 a nie -2
				cell = row.AddCell()
				cell.SetStyle(sumStyle)
				cell = row.AddCell()
				cell.SetStyle(sumStyle)
				cell = row.AddCell()
				cell.SetStyle(sumStyle)
				cell = row.AddCell()
				cell.SetStyle(sumStyle)
				//cell.Value = xlsx.GetCellIDStringFromCoords(6, rowNbr-2-rowInMonth+1) //rowInMont+1 bo zaczynamy liczyc od komorki w ktorej jestesmy
				cell = row.AddCell()
				cell.SetStyle(sumStyle)
				cell = row.AddCell()
				cell.SetStyle(sumStyle)
				//cell.Value = cellCoordinate
				formule := "sum(" + xlsx.GetCellIDStringFromCoords(7, rowNbr-2-rowInMonth+1) + ":" + cellCoordinate + ")"
				cell.SetFormula(formule)
				cell.NumFmt = "##,##0.00"
				cell = row.AddCell()
				cell.SetStyle(sumStyle)
				cell = row.AddCell()
				cell.SetStyle(sumStyle)
				cell = row.AddCell()
				cell.SetStyle(sumStyleRight)
				rowInMonth = 1
				rowNbr++

			}
			if currnetMonth == "" {
				currnetMonth = month
				//currentPeriod = strings.Split(wiersze[i].opis, ",")[1]
				//month = strconv.Itoa(wiersze[i].accountingMonth)
				currentPeriod = strconv.Itoa(wiersze[i].accountingMonth) + "/" + strconv.Itoa(wiersze[i].accountingYear)
				rowNbr++
			}

			row = sheet.AddRow()
			cell = row.AddCell()
			cell = row.AddCell()
			cell.SetStyle(col0StyleLeft)
			cell = row.AddCell()
			cell.Value = wiersze[i].nrDok
			cell.SetStyle(defCellStyle)
			cell = row.AddCell()

			cell.SetDate(wiersze[i].data)

			cell.SetStyle(defCellStyle)
			cell = row.AddCell()
			cell.Value = wiersze[i].referencja
			cell.SetStyle(defCellStyle)
			cell = row.AddCell()
			cell.Value = wiersze[i].opis
			cell.SetStyle(defCellStyle)
			cell = row.AddCell()
			if wiersze[i].ilosc > 0 {
				cell.SetFloatWithFormat(wiersze[i].ilosc, "##,##")
				cell.SetStyle(defCellStyle)
			} else {
				cell.SetStyle(warning)
			}
			cell = row.AddCell()
			cell.SetFloatWithFormat(wiersze[i].kwotaWkr, "##,##0.00")
			cell.SetStyle(defCellStyle)
			cell = row.AddCell()
			if wiersze[i].ilosc > 0 {
				cell.SetFloatWithFormat(wiersze[i].kwotaWkr/wiersze[i].ilosc, "0.00")
			}
			cell.SetStyle(defCellStyle)
			cell = row.AddCell()
			cell.Value = wiersze[i].przec
			cell.SetStyle(defCellStyle)
			cell = row.AddCell()
			cell.Value = "TBD"
			cell.SetStyle(col0StyleRight)
			//k := len(wiersze)
			// fmt.Println(k)
			// fmt.Println(i)

			//debug purposes
			cell = row.AddCell()
			cell.Value = month

			cell = row.AddCell()
			cell.Value = month

			cell = row.AddCell()
			cell.SetValue(wiersze[i].accountingMonth)

			totalkWhAll += wiersze[i].ilosc
			totalPLNAll += wiersze[i].kwotaWkr
		}
	}
	//podsumowanie ostatni miesiac
	rowNbr++
	row = sheet.AddRow()
	cell = row.AddCell()
	cell = row.AddCell()
	cell.Value = currentPeriod
	cell.SetStyle(sumStyleLeft)
	currentPeriod = strings.Split(wiersze[v].opis, ",")[1]
	cellCoordinate := xlsx.GetCellIDStringFromCoords(7, rowNbr-3) //wykonywane w wierszu ponizej podsumowania (inny miesiac) + wiersz podsumowujacy, wiersze i kolumny numerwoane od 0 w go dlatego -3 a nie -2
	cell = row.AddCell()
	cell.SetStyle(sumStyle)
	cell = row.AddCell()
	cell.SetStyle(sumStyle)
	cell = row.AddCell()
	cell.SetStyle(sumStyle)
	cell = row.AddCell()
	cell.SetStyle(sumStyle)
	//cell.Value = xlsx.GetCellIDStringFromCoords(6, rowNbr-2-rowInMonth) //rowInMont+1 bo zaczynamy liczyc od komorki w ktorej jestesmy
	cell = row.AddCell()
	cell.SetStyle(sumStyle)
	cell = row.AddCell()
	cell.SetStyle(sumStyle)
	//cell.Value = cellCoordinate
	formule := "sum(" + xlsx.GetCellIDStringFromCoords(7, rowNbr-2-rowInMonth) + ":" + cellCoordinate + ")"
	cell.SetFormula(formule)
	cell.NumFmt = "##,##0.00"
	cell = row.AddCell()
	cell.SetStyle(sumStyle)
	cell = row.AddCell()
	cell.SetStyle(sumStyle)
	cell = row.AddCell()
	cell.SetStyle(sumStyleRight)
	rowInMonth = 1
	rowNbr++

	//podsumowanie total
	row = sheet.AddRow()
	cell = row.AddCell()
	cell = row.AddCell()
	cell.SetStyle(col0StyleLeft)
	cell = row.AddCell()
	cell.SetStyle(defCellStyle)
	cell = row.AddCell()
	cell.SetStyle(defCellStyle)
	cell = row.AddCell()
	cell.SetStyle(defCellStyle)
	cell = row.AddCell()
	cell.SetStyle(defCellStyle)
	cell = row.AddCell()
	cell.SetStyle(defCellStyle)
	cell = row.AddCell()
	cell.SetStyle(defCellStyle)
	cell = row.AddCell()
	cell.SetStyle(defCellStyle)
	cell = row.AddCell()
	cell.SetStyle(defCellStyle)
	cell = row.AddCell()
	cell.SetStyle(col0StyleRight)
	row = sheet.AddRow()
	row.AddCell()
	cell = row.AddCell()
	cell.Value = "Total"
	cell.SetStyle(sumTotalStyleLeft)
	//
	cell = row.AddCell()
	cell.SetStyle(sumTotalStyle)
	cell = row.AddCell()
	cell.SetStyle(sumTotalStyle)
	cell = row.AddCell()
	cell.SetStyle(sumTotalStyle)
	cell = row.AddCell()
	cell.SetStyle(sumTotalStyle)
	cell = row.AddCell()
	cell.SetStyle(sumTotalStyle)
	cell.SetFloatWithFormat(totalkWhAll, "##,##")
	cell = row.AddCell()
	cell.SetStyle(sumTotalStyle)
	cell.SetFloatWithFormat(totalPLNAll, "##,##0.00")
	cell = row.AddCell()
	cell.SetStyle(sumTotalStyle)
	cell = row.AddCell()
	cell.SetStyle(sumTotalStyle)
	cell = row.AddCell()
	cell.SetStyle(sumTotalStyleRight)
	return fileObr
}

func setIloscPerObrotNew(wiersze []energia) {
	month := 0
	double := false
	for i := range wiersze {
		if wiersze[i].referencja == "97971/1803/00168" && wiersze[i].zaklad == 3437 {
			fmt.Println(wiersze[i])
		}
		currentMonth := wiersze[i].accountingMonth
		if wiersze[i].przypisanie == "obrót" {
			if wiersze[i].fakOkresowa == false {
				object := strings.Split(wiersze[i].opis, ",")[2] //czego dotyczy: zaklad, portiernia, etc.
				iloscTemp := 0.0
				if month != currentMonth {
					for j := range wiersze {
						//dla faktury dystrybucyjnej dla tego samego obiektu oraz miesiaca i roku ksiegowania
						if wiersze[j].przypisanie == "dystrybucja" && strings.Split(wiersze[j].opis, ",")[2] == object && wiersze[i].accountingMonth == wiersze[j].accountingMonth && wiersze[i].accountingYear == wiersze[j].accountingYear {
							iloscTemp += wiersze[j].ilosc //suma ilosc w razie wystapienia korekty
						}
					}
					double = false
					month = currentMonth
				} else {
					double = true
					iloscTemp = 0.0 //unikamy sumowania ilosci dla drugiej faktury wystawionej dla tego samego obiektu w tym samym miesiacu - WYMAGANA KOREKTA RECZNA PO UTWORZENIU RAPORTU
				}
				if wiersze[i].ilosc < iloscTemp && wiersze[i].ilosc > 0 {
					wiersze[i].ilosc = iloscTemp
				} else if wiersze[i].ilosc < 0 {
					wiersze[i].ilosc += iloscTemp
				}
				if double == true {
					wiersze[i].ilosc = 0.0
				}
			} else if wiersze[i].fakOkresowa == true {
				startMonth := int(wiersze[i].dateFrom.Month())
				object := strings.Split(wiersze[i].opis, ",")[2] //czego dotyczy: zaklad, portiernia, etc.
				iloscTemp := 0.0
				if month != currentMonth {
					for j := range wiersze {
						//dla faktury dystrybucyjnej dla tego samego obiektu oraz miesiaca i roku ksiegowania
						if wiersze[j].przypisanie == "dystrybucja" && strings.Split(wiersze[j].opis, ",")[2] == object && wiersze[i].accountingYear == wiersze[j].accountingYear && startMonth <= wiersze[j].accountingMonth && wiersze[i].accountingMonth >= wiersze[j].accountingMonth {
							iloscTemp += wiersze[j].ilosc //suma ilosc w razie wystapienia korekty
						}
					}
					double = false
					month = currentMonth
				} else {
					double = true
					iloscTemp = 0.0 //unikamy sumowania ilosci dla drugiej faktury wystawionej dla tego samego obiektu w tym samym miesiacu - WYMAGANA KOREKTA RECZNA PO UTWORZENIU RAPORTU
				}
				if wiersze[i].ilosc < iloscTemp && wiersze[i].ilosc > 0 {
					wiersze[i].ilosc = iloscTemp
				} else if wiersze[i].ilosc < 0 {
					wiersze[i].ilosc += iloscTemp
				}
				if double == true {
					wiersze[i].ilosc = 0.0
				}
			}
		}
	}
}

func setIloscPerObrot(wiersze []energia) {

	var month string
	var year string
	var monthJ string
	var yearJ string

	for i := range wiersze {
		if wiersze[i].przypisanie == "obrót" {
			object := strings.Split(wiersze[i].opis, ",")[2] //czego dotyczy: zaklad, portiernia, etc.
			//period := strings.Split(wiersze[i].opis, ",")[1] //okres faktury z opisu

			monthArr := strings.Split((strings.Split(wiersze[i].opis, ",")[1]), "/")
			if len(monthArr) < 3 {
				month = monthArr[0]
				year = monthArr[1]
			} else {
				month = monthArr[3]
				year = monthArr[4]
			}

			iloscTemp := 0.0
			for j := range wiersze {

				monthArrJ := strings.Split((strings.Split(wiersze[j].opis, ",")[1]), "/")
				if len(monthArrJ) < 3 {
					monthJ = monthArrJ[0]
					yearJ = monthArrJ[1]
				} else {
					monthJ = monthArrJ[3]

					fmt.Println(i)
					fmt.Println(wiersze[i])
					fmt.Println(monthArrJ)

					yearJ = monthArrJ[4]

				}

				if wiersze[j].przypisanie == "dystrybucja" && strings.Split(wiersze[j].opis, ",")[2] == object && month == monthJ && year == yearJ {
					iloscTemp += wiersze[j].ilosc //suma ilosc w razie wystapienia korekty
				}
			}
			if wiersze[i].ilosc < iloscTemp && wiersze[i].ilosc > 0 {
				wiersze[i].ilosc = iloscTemp
			} else if wiersze[i].ilosc < 0 {
				wiersze[i].ilosc += iloscTemp
			}
		}

	}
}

func main() {

	const shortForm = "2018-Jan-01"

	excelFileName := "C:\\tmp\\175300_3.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Printf("error")
	}

	endOfRows = 0
	wiersze := make([]energia, 0)
	sheet := xlFile.Sheets[0]
	j := 7 //rows index, 7 - where the data in xlsx file starts
	for endOfRows == 0 {
		i := 2 //columns iteration, i - columns index, 2 -where the data in xlsx file starts Cell (C8)

		wiersz.przec = sheet.Cell(j, i).String()
		wiersz.zaklad, _ = sheet.Cell(j, i+1).Int()
		wiersz.po = sheet.Cell(j, i+2).String()

		temp := strings.Replace(sheet.Cell(j, i+4).String(), ".", "-", -1)
		t, _ := time.Parse("02-01-2006", temp)

		wiersz.data = t
		wiersz.przypisanie = sheet.Cell(j, i+6).String()
		wiersz.nrDok = sheet.Cell(j, i+7).String()
		wiersz.rodzaj = sheet.Cell(j, i+8).String()

		temp2 := strings.Replace(sheet.Cell(j, i+9).String(), ".", "-", -1)
		t2, _ := time.Parse("02-01-2006", temp2)

		//wiersz.dataDokumentu = sheet.Cell(j, i+9).String()
		wiersz.dataDokumentu = t2
		wiersz.referencja = sheet.Cell(j, i+10).String()
		wiersz.kk, _ = sheet.Cell(j, i+11).Int()

		kwotaWkr := sheet.Cell(j, i+12).String()
		kwotaWkr = strings.Replace(kwotaWkr, ".", "", 1)
		kwotaWkr = strings.TrimLeft(kwotaWkr, " ")
		kwotaWkr = strings.Replace(kwotaWkr, ",", ".", 1)
		kwotaWkrFloat, _ := strconv.ParseFloat(kwotaWkr, 64)

		wiersz.kwotaWkr = kwotaWkrFloat
		wiersz.walKr = sheet.Cell(j, i+13).String()

		ilosc := sheet.Cell(j, i+14).String()

		if sheet.Cell(j, i+16).String() == "roz.energii,03/2018,zakład,3437,dystrybucja" {
			fmt.Println(ilosc)
		}
		if ilosc == "1.008.931,0" {
			fmt.Println(ilosc)
		}

		ilosc = strings.Replace(ilosc, ".", "", -1)
		ilosc = strings.TrimLeft(ilosc, " ")
		ilosc = strings.Replace(ilosc, ",", ".", 1)
		iloscFloat, _ := strconv.ParseFloat(ilosc, 64)

		wiersz.ilosc = iloscFloat
		wiersz.pd = sheet.Cell(j, i+15).String()
		wiersz.opis = sheet.Cell(j, i+16).String()

		if wiersz.przec != "" {
			wiersze = append(wiersze, wiersz)
		}

		j++
		cell := sheet.Cell(j, 1)
		if cell.String() == "**" {
			endOfRows = 1
		}
	}

	// for i := range wiersze {
	// 	fmt.Println(wiersze[i])
	// 	fmt.Printf("\n")

	// }

	sort.Slice(wiersze, func(i, j int) bool {
		return wiersze[i].zaklad < wiersze[j].zaklad
	})

	slicePerPlant, v := splitSlicePerPlant(wiersze)
	// for i := 0; i < v+1; i++ {
	// 	fmt.Println("PLANT: " + strconv.Itoa(slicePerPlant[i][0].zaklad))
	// 	for j := range slicePerPlant[i] {
	// 		fmt.Println(slicePerPlant[i][j])
	// 		fmt.Printf("\n")
	// 	}
	// }

	fileXLSX := *xlsx.NewFile()
	fileXLSX2 := *xlsx.NewFile()
	fileXLSX3 := *xlsx.NewFile()
	for z := 0; z < v+1; z++ {

		//sortowanie po dacie per zaklad

		sort.Slice(slicePerPlant[z], func(i, j int) bool {
			return slicePerPlant[z][i].data.Before(slicePerPlant[z][j].data)
		})

		slicePerPlant[z] = invoiceAggregation(slicePerPlant[z])
		slicePerPlant[z] = checkInvoiceTimePeriod(slicePerPlant[z])
		slicePerPlant[z] = setDateFromDateTo(slicePerPlant[z])
		slicePerPlant[z] = setAccountingMonth(slicePerPlant[z])
		//setIloscPerObrot(slicePerPlant[z])
		setIloscPerObrotNew(slicePerPlant[z])

		fileXLSX = saveToNewSheet(slicePerPlant[z], "obrót", fileXLSX)
		fileXLSX2 = saveToNewSheet(slicePerPlant[z], "dystrybucja", fileXLSX2)
		fileXLSX3 = saveToNewSheet(slicePerPlant[z], "obrót+dystrybucja", fileXLSX3)

	}
	err = fileXLSX.Save("C:\\tmp\\test-obr.xlsx")
	err = fileXLSX2.Save("C:\\tmp\\test-dys.xlsx")
	err = fileXLSX3.Save("C:\\tmp\\test-dysobr.xlsx")

}
