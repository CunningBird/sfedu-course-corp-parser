package com.cunningbird.sfedu.corp

import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow

data class UpdSheet(
    var invoice: XSSFCell?,
    var date: XSSFCell?,
    var listOfRows: List<XSSFRow>?
)