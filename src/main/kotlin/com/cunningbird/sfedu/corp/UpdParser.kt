package com.cunningbird.sfedu.corp

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.io.IOException

class UpdParser(
    private var file: File,
    private var xssfWorkbook: XSSFWorkbook,
    private var mapOfSpecialCellIndexes: HashMap<String, Pair<Int, Int>>
) {

    fun getDataFromBook(): List<UpdSheet> {
        val quantityOfSheets = xssfWorkbook.numberOfSheets
        var i = 0
        return List(quantityOfSheets) { getDataFromSheet(i++) }
    }

    private fun getDataFromSheet(number: Int) =
        UpdSheet(getInvoiceFromSheet(number), getDateFromSheet(number), getRowsOfGoods(number))

    private fun getInvoiceFromSheet(number: Int) =
        xssfWorkbook.getSheetAt(number).getRow(mapOfSpecialCellIndexes["invoice cell"]!!.first)
            .getCell(mapOfSpecialCellIndexes["invoice cell"]!!.second)

    private fun getDateFromSheet(number: Int) =
        xssfWorkbook.getSheetAt(number).getRow(mapOfSpecialCellIndexes["date cell"]!!.first)
            .getCell(mapOfSpecialCellIndexes["date cell"]!!.second)

    private fun getRowsOfGoods(number: Int): MutableList<XSSFRow> {
        var currentRow = xssfWorkbook.getSheetAt(number).getRow(mapOfSpecialCellIndexes["rows of goods"]!!.first)
        var i = 0
        val listOfRows = mutableListOf<XSSFRow>()
        while (currentRow.getCell(0).stringCellValue.isNotEmpty()) {
            listOfRows.add(currentRow)
            ++i
            currentRow = xssfWorkbook.getSheetAt(number).getRow(mapOfSpecialCellIndexes["rows of goods"]!!.first + i)
        }
        return listOfRows
    }

    fun createReportTable(listUpd: List<UpdSheet>) {
        createSheetTemplate(0)
        inputUpdsInTemplate(listUpd)
        fancyTable()
    }

    private fun fancyTable() {
        val sheet = xssfWorkbook.getSheet("Отчет")
        for (i in 0..20) sheet.autoSizeColumn(i, true)
    }

    private fun inputUpdsInTemplate(listUpd: List<UpdSheet>) {
        val sheet = xssfWorkbook.getSheet("Отчет")
        var i = 2
        listUpd.forEach {
            it.listOfRows?.forEach { it2 ->
                var j = 2
                val row = sheet.createRow(++i)
                row.createCell(0).setCellValue(it.invoice?.stringCellValue)
                row.createCell(1).setCellValue(it.date?.stringCellValue)
                it2.forEachIndexed { ind, _ ->
                    val cell = it2.getCell(ind).apply { cellType = CellType.STRING }
                    if (cell.stringCellValue.isNotEmpty()) {
                        row.createCell(j++).setCellValue(cell.stringCellValue)
                    }
                }
            }
        }
        createOutputStream()
    }

    private fun createSheetTemplate(number: Int) {
        var i = 0
        val listOfRows = MutableList<XSSFRow>(mapOfSpecialCellIndexes["data rows and their quantity"]!!.second) {
            xssfWorkbook.getSheetAt(number).getRow(mapOfSpecialCellIndexes["data rows and their quantity"]!!.first + i++)
        }
        val newSheet = xssfWorkbook.createSheet("Отчет").apply {
            createRow(0)
            createRow(1)
            createRow(2)
        }

        newSheet.addMergedRegion(CellRangeAddress(0, 1, 0, 0))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 1, 1))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 2, 2))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 3, 3))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 4, 4))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 5, 5))
        newSheet.addMergedRegion(CellRangeAddress(0, 0, 6, 7))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 8, 8))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 9, 9))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 10, 10))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 11, 11))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 12, 12))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 13, 13))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 14, 14))
        newSheet.addMergedRegion(CellRangeAddress(0, 0, 15, 16))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 17, 17))
        newSheet.addMergedRegion(CellRangeAddress(0, 0, 18, 19))
        newSheet.addMergedRegion(CellRangeAddress(0, 1, 20, 20))

        newSheet.getRow(0).createCell(0).setCellValue("Счет-фактура №")
        newSheet.getRow(0).createCell(1).setCellValue("Дата")
        newSheet.getRow(0).createCell(2).setCellValue("Код товара/ работ, услуг")
        newSheet.getRow(0).createCell(3).setCellValue("№ п/п")
        newSheet.getRow(0).createCell(4)
            .setCellValue("Наименование товара (описание выполненных работ, оказанных услуг), имущественного права")
        newSheet.getRow(0).createCell(5).setCellValue("Код вида товара")
        newSheet.getRow(0).createCell(6).setCellValue("Единица измерения")
        newSheet.getRow(1).createCell(6).setCellValue("код")
        newSheet.getRow(1).createCell(7).setCellValue("условное обозначение (национальное)")
        newSheet.getRow(0).createCell(8).setCellValue("Количество (объем)")
        newSheet.getRow(0).createCell(9).setCellValue("Цена (тариф) за единицу измерения")
        newSheet.getRow(0).createCell(10)
            .setCellValue("Стоимость товаров (работ, услуг), имущественных прав без налога — всего")
        newSheet.getRow(0).createCell(11).setCellValue("В том числе сумма акциза")
        newSheet.getRow(0).createCell(12).setCellValue("Налоговая ставка")
        newSheet.getRow(0).createCell(13).setCellValue("Сумма налога, предъявляемая покупателю")
        newSheet.getRow(0).createCell(14)
            .setCellValue("Стоимость товаров (работ, услуг), имущественных прав с налогом — всего")
        newSheet.getRow(0).createCell(15).setCellValue("Страна происхождения товара")
        newSheet.getRow(1).createCell(15).setCellValue("Цифровой код")
        newSheet.getRow(1).createCell(16).setCellValue("Краткое наименование")
        newSheet.getRow(0).createCell(16)
            .setCellValue("Регистрационный номер декларации на товары или регистрационный номер партии товара, подлежащего прослеживаемости")
        newSheet.getRow(0).createCell(17)
            .setCellValue("Регистрационный номер декларации на товары или регистрационный номер партии товара, подлежащего прослеживаемости")
        newSheet.getRow(0).createCell(18)
            .setCellValue("Количественная единица измерения товара, используемая в целях прослеживаемости")
        newSheet.getRow(1).createCell(18).setCellValue("код")
        newSheet.getRow(1).createCell(19).setCellValue("условное обозначение")
        newSheet.getRow(0).createCell(19)
            .setCellValue("Кол.товара, подлежащего прослеживаемости, в количественной ед.изм. Товара")

        newSheet.getRow(2).createCell(0).setCellValue("А")
        newSheet.getRow(2).createCell(1).setCellValue("1")
        newSheet.getRow(2).createCell(2).setCellValue("1а")
        newSheet.getRow(2).createCell(3).setCellValue("1б")
        newSheet.getRow(2).createCell(4).setCellValue("2")
        newSheet.getRow(2).createCell(6).setCellValue("2а")
        newSheet.getRow(2).createCell(7).setCellValue("3")
        newSheet.getRow(2).createCell(8).setCellValue("4")
        newSheet.getRow(2).createCell(9).setCellValue("5")
        newSheet.getRow(2).createCell(10).setCellValue("6")
        newSheet.getRow(2).createCell(11).setCellValue("7")
        newSheet.getRow(2).createCell(12).setCellValue("8")
        newSheet.getRow(2).createCell(13).setCellValue("9")
        newSheet.getRow(2).createCell(14).setCellValue("10")
        newSheet.getRow(2).createCell(15).setCellValue("10а")
        newSheet.getRow(2).createCell(16).setCellValue("11")
        newSheet.getRow(2).createCell(17).setCellValue("12")
        newSheet.getRow(2).createCell(18).setCellValue("12а")
        newSheet.getRow(2).createCell(19).setCellValue("13")
        createOutputStream()
    }

    private fun createOutputStream() {
        var fos: FileOutputStream? = null
        try {
            fos = FileOutputStream(file)
            xssfWorkbook.write(fos)
        } catch (e: IOException) {
            e.printStackTrace()
        } finally {
            if (fos != null) {
                try {
                    fos.flush()
                    fos.close()
                } catch (e: IOException) {
                    e.printStackTrace()
                }
            }
        }
    }
}