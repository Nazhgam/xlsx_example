package main

import (
	"fmt"
	"log"

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
	// Открываем существующий файл Excel.
	filePath := "C:/Users/kirig/wildberries/experiments/xlsx_example/Book1.xlsx"
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Ошибка при открытии файла: %v", err)
	}

	// Проверяем, был ли файл успешно открыт.
	if f == nil {
		log.Fatal("Файл не удалось открыть")
	}

	datas := getMockData()

	for ind, data := range datas {
		// Устанавливаем новые значения для определенных ячеек.
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

	// Сохраняем изменения в файле.
	if err = f.Save(); err != nil {
		log.Fatalf("О��ибка при сохранении файла: %v", err)
	}

	fmt.Println("Значения успешно обновлены в файле Excel.")
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
