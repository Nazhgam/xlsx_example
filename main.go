package main

import (
	"fmt"
	"log"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"
)

type XlsxDatas struct {
	Date           string
	DocNum         string
	BIKPayer       string
	PayerBank      string
	PayerName      string
	INNPayer       string
	PayerNumAkk    string
	ReceiverBIK    string
	ReceiverBank   string
	ReceiverName   string
	ReveiverINN    string
	ReceiverAkkNum string
	DebuteSum      float64
	KreditSum      float64
	SaldoAfterOper float64
	PaymentName    string
}

func main() {
	// Создаем новый файл Excel
	f := excelize.NewFile()

	// Заполняем первую строку колонкой A
	f.SetCellValue("Sheet1", "A2", `ООО "Вайлдберриз Банк"`)
	f.SetCellValue("Sheet1", "A3", `Номер счета и наименование клиента: 40702810700000000321 Общество с ограниченной ответственностью "ВБ Восток"`)
	f.SetCellValue("Sheet1", "A4", `Период выгрузки с:  01-01-2024 - 30-03-2024`)
	f.SetCellValue("Sheet1", "A6", `Входящий остаток (в валюте счета): `)

	// Добавляем данные в таблицу начиная с 7 строки
	datas := getMockData()

	startRow := 7

	// Добавляем заголовки таблицы
	for col, header := range map[string]string{
		"A": "Дата",
		"B": "Номер документа",
		"C": "БИК Банка плательщика",
		"D": "Банк плательщика",
		"E": "Наименование плательщика",
		"F": "ИНН плательщика",
		"G": "№ счета плательщика",
		"H": "БИК банка получателя",
		"I": "Банк получателя",
		"J": "Наименование получателя",
		"K": "ИНН получателя",
		"L": "№ счета получателя",
		"M": "Сумма операции по дебету счета",
		"N": "Сумма операции по кредиту счета",
		"O": "Сальдо после операции",
		"P": "Назначение платежа",
	} {
		f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", col, startRow), header)

	}

	// Добавляем данные в таблицу
	for ind, data := range datas {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", 8+ind), data.Date)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", 8+ind), data.DocNum)
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", 8+ind), data.BIKPayer)
		f.SetCellValue("Sheet1", fmt.Sprintf("D%d", 8+ind), data.PayerBank)
		f.SetCellValue("Sheet1", fmt.Sprintf("E%d", 8+ind), data.PayerName)
		f.SetCellValue("Sheet1", fmt.Sprintf("F%d", 8+ind), data.INNPayer)
		f.SetCellValue("Sheet1", fmt.Sprintf("G%d", 8+ind), data.PayerNumAkk)
		f.SetCellValue("Sheet1", fmt.Sprintf("H%d", 8+ind), data.ReceiverBIK)
		f.SetCellValue("Sheet1", fmt.Sprintf("I%d", 8+ind), data.ReceiverBank)
		f.SetCellValue("Sheet1", fmt.Sprintf("J%d", 8+ind), data.ReceiverName)
		f.SetCellValue("Sheet1", fmt.Sprintf("K%d", 8+ind), data.ReveiverINN)
		f.SetCellValue("Sheet1", fmt.Sprintf("L%d", 8+ind), data.ReceiverAkkNum)
		f.SetCellValue("Sheet1", fmt.Sprintf("M%d", 8+ind), data.DebuteSum)
		f.SetCellValue("Sheet1", fmt.Sprintf("N%d", 8+ind), data.KreditSum)
		f.SetCellValue("Sheet1", fmt.Sprintf("O%d", 8+ind), data.SaldoAfterOper)
		f.SetCellValue("Sheet1", fmt.Sprintf("P%d", 8+ind), data.PaymentName)

	}
	setBorder(f, len(datas))
	// Сохраняем файл
	if err := f.SaveAs("Book0.xlsx"); err != nil {
		log.Fatalf("Ошибка при сохранении файла: %v", err)
	}
	err := getMaxColWidth(f, len(datas))
	if err != nil {
		log.Fatalf("Ошибка при настройке ШИРИНЕ файла: %v", err)
	}
	// Сохраняем файл
	if err := f.SaveAs("Book0.xlsx"); err != nil {
		log.Fatalf("Ошибка при сохранении файла: %v", err)
	}
}

func setBorder(f *excelize.File, lastIndex int) {
	// Создаем первый стиль для диапазона A7:P7
	style1, err := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		fmt.Println(err)
		return
	}

	// Создаем второй стиль для диапазона A8:P16
	style2, err := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		fmt.Println(err)
		return
	}

	// Создаем третий стиль для диапазона A17:P17
	style3, err := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		fmt.Println(err)
		return
	}
	f.SetCellStyle("Sheet1", "A7", fmt.Sprintf("O%d", 7+lastIndex-1), style1)
	f.SetCellStyle("Sheet1", fmt.Sprintf("P%d", 7), fmt.Sprintf("P%d", 7+lastIndex-1), style2)
	f.SetCellStyle("Sheet1", fmt.Sprintf("A%d", 7+lastIndex), fmt.Sprintf("P%d", 7+lastIndex), style3)
}

func getMockData() []XlsxDatas {
	var res = make([]XlsxDatas, 10)

	for ind := range res {
		res[ind].Date = fmt.Sprintf("2024-07-%d", ind)
		res[ind].DocNum = fmt.Sprintf("doc num: %d", ind)
		res[ind].BIKPayer = fmt.Sprintf("BIK payer: %d", ind)
		res[ind].PayerBank = fmt.Sprintf("Bank payer: %d", ind)
		res[ind].PayerName = fmt.Sprintf("payer name: %d", ind)
		res[ind].INNPayer = fmt.Sprintf("inn payer: %d", ind)
		res[ind].PayerNumAkk = fmt.Sprintf("payer num akk: %d", ind)
		res[ind].ReceiverBIK = fmt.Sprintf("receiver BIK: %d", ind)
		res[ind].ReceiverBank = fmt.Sprintf("receiver bank: %d", ind)
		res[ind].ReceiverName = fmt.Sprintf("receiver name: %d", ind)
		res[ind].ReveiverINN = fmt.Sprintf("reveiver inn: %d", ind)
		res[ind].ReceiverAkkNum = fmt.Sprintf("reveiver akk num: %d", ind)
		res[ind].DebuteSum = 10.5 + float64(ind)
		res[ind].KreditSum = 10.5 + float64(ind)
		res[ind].SaldoAfterOper = 10.5 + float64(ind)
		res[ind].PaymentName = fmt.Sprintf("payment name: %d", ind)
	}

	return res
}

func getMaxColWidth(f *excelize.File, endRow int) error {
	for col := 'A'; col <= 'P'; col++ {
		maxWidth := 0
		for ind := 7; ind <= 7+endRow; ind++ {
			val, err := f.GetCellValue("Sheet1", fmt.Sprintf("%c%d", col, ind))
			if err != nil {
				return err
			}
			if len(val) > maxWidth {
				maxWidth = utf8.RuneCountInString(val)
			}

		}
		err := f.SetColWidth("Sheet1", string(col), string(col), float64(maxWidth))
		if err != nil {
			return err
		}
	}

	return nil
}
