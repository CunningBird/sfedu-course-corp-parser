package com.cunningbird.sfedu.corp

import org.apache.poi.sl.usermodel.ObjectMetaData
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.util.*


fun main(args: Array<String>) {

//    val file = File("C:\\Users\\serj-\\Desktop\\upd2.xlsx")
//    val fis = FileInputStream(file)

    val excelFile = File(Objects.requireNonNull(ObjectMetaData.Application::class.java.classLoader.getResource("upd2.xlsx")).file)
    val excelFileIs = FileInputStream(excelFile)
    val book = XSSFWorkbook(excelFileIs)

    val hashMapOfSpecialIndexes = hashMapOf(
        "invoice cell" to Pair(1, 23),
        "date cell" to Pair(1, 33),
        "rows of goods" to Pair(22, 0),
        "data rows and their quantity" to Pair(17, 5)
    )

    val updParser = UpdParser(excelFile, book, hashMapOfSpecialIndexes)
    val data = updParser.getDataFromBook()
    updParser.createReportTable(data)
    print("success")
}

