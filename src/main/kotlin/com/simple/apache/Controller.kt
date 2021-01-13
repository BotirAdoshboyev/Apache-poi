package com.simple.apache

import org.apache.poi.ss.util.CellUtil.createCell
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.http.*
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RestController
import java.io.ByteArrayOutputStream
import java.io.OutputStream
import java.util.*

data class Dto(
    var id: Long,
    var name: String,
    var total: Long,
    var new: Long,
    var ready: Long,
    var issue: Long,
    var print: Long
)


fun getItems(): List<Dto>{

    val res = mutableListOf<Dto>()

    for (i in 0..30){
        val item = Dto(1, "Bdlewkmdkwednwjndotir", 250, 4000, 253, 500, 265)
        res.add(i, item)
    }
    return res
}
@RestController
class Controller {
    @PostMapping
    fun post() : ResponseEntity<ByteArray>{
        var rowIndex = 1
        val list = getItems()

        val xlWB = XSSFWorkbook()
        val xlWS = xlWB.createSheet("statistika")

        val boldCellStyle: XSSFCellStyle = xlWB.createCellStyle()
        val isBold: XSSFFont = xlWB.createFont()

        isBold.bold = true
        boldCellStyle.setFont(isBold)

        val headerRow = xlWS.createRow(0)

        val headerIdCell = createCell(headerRow, 0, "Id")
        headerIdCell.cellStyle = boldCellStyle
        val headerNameCell = createCell(headerRow, 1, "Viloyat")
        headerNameCell.cellStyle = boldCellStyle
        val headerTotalCell = createCell(headerRow, 2, "Umumiy")
        headerTotalCell.cellStyle = boldCellStyle
        val headerNewCell = createCell(headerRow,3, "Yangi")
        headerNewCell.cellStyle = boldCellStyle
        val headerReadyCell = createCell(headerRow, 4, "Tayyor")
        headerReadyCell.cellStyle = boldCellStyle
        val headerIssueCell = createCell(headerRow, 5, "Berilgan")
        headerIssueCell.cellStyle = boldCellStyle
        val headerPrintCell = createCell(headerRow, 6, "Pechat qilingan")
        headerPrintCell.cellStyle = boldCellStyle

        var totalOfTotal: Long = 0
        var totalOfNew: Long = 0
        var totalOfReady: Long = 0
        var totalOfIssue: Long = 0
        var totalOfPrint: Long = 0

        list.forEach{
            totalOfTotal += it.total
            totalOfNew += it.new
            totalOfReady += it.ready
            totalOfIssue += it.issue
            totalOfPrint += it.print

            val xlRow = xlWS.createRow(rowIndex++)

            val idCell = xlRow.createCell(0)
            idCell.setCellValue(it.id.toString())

            val nameCell = xlRow.createCell(1)
            nameCell.setCellValue(it.name)

            val totalCell = xlRow.createCell(2)
            totalCell.setCellValue(it.total.toString())

            val newCell = xlRow.createCell(3)
            newCell.setCellValue(it.new.toString())

            val readyCell = xlRow.createCell(4)
            readyCell.setCellValue(it.ready.toString())

            val issueCell = xlRow.createCell(5)
            issueCell.setCellValue(it.issue.toString())

            val printCell = xlRow.createCell(6)
            printCell.setCellValue(it.print.toString())
        }
        val footerRow = xlWS.createRow(rowIndex)
        val footerTextCell = createCell(footerRow, 1, "Umumiy")
        footerTextCell.cellStyle = boldCellStyle
        val footerTotalCell = createCell(footerRow, 2, totalOfTotal.toString())
        footerTotalCell.cellStyle = boldCellStyle
        val footerNewCell = createCell(footerRow, 3, totalOfNew.toString())
        footerNewCell.cellStyle = boldCellStyle
        val footerReadyCell = createCell(footerRow, 4, totalOfReady.toString())
        footerReadyCell.cellStyle = boldCellStyle
        val footerIssueCell = createCell(footerRow, 5, totalOfIssue.toString())
        footerIssueCell.cellStyle = boldCellStyle
        val footerPrintCell = createCell(footerRow, 6, totalOfPrint.toString())
        footerPrintCell.cellStyle = boldCellStyle

//        for (i in 0..6){
//            xlWS.autoSizeColumn(i)
//        }

        xlWS.autoSizeColumn(0)
        xlWS.autoSizeColumn(1)
        xlWS.autoSizeColumn(2)
        xlWS.autoSizeColumn(3)
        xlWS.autoSizeColumn(4)
        xlWS.autoSizeColumn(5)
        xlWS.autoSizeColumn(6)
        xlWS.createFreezePane(0, 1)
        val byteArrayOS = ByteArrayOutputStream()
        xlWB.write(byteArrayOS)
        xlWB.close()


        val headers = HttpHeaders().apply {
            contentType = MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8")
            contentDisposition = ContentDisposition.builder("inline").filename("test.xlsx").build()
            contentLength = byteArrayOS.size().toLong()
        }

        return ResponseEntity(byteArrayOS.toByteArray(), headers, HttpStatus.OK)
    }
}