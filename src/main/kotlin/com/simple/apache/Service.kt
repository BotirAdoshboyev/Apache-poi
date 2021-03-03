package com.simple.apache

import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.http.*
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestBody
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RestController
import java.io.ByteArrayOutputStream
import java.text.SimpleDateFormat
import java.time.Instant
import java.time.LocalDate
import java.time.ZoneId
import java.time.ZoneOffset
import java.util.*


data class Dto(
    var id: Long,
    var name: String,
    var total: Long,
    var type: Long,
    var date: Long
)

data class IntervalDto(
    val startDate: Long,
    val endDate: Long
)

fun getItems(date: Long, type: Long): List<Dto>{
    val res = mutableListOf<Dto>()
    var k = 0
    for (i in 0..10){
        k += 4
        val item = Dto(i.toLong(), "Samarqand $i", 20+i.toLong() + k.toLong(), type, date)
        res.add(i, item)
    }
    return res
}
@RestController
@RequestMapping
class Controller {
    @PostMapping
    fun post() : ResponseEntity<ByteArray>{
        var totalInRegions = mutableMapOf<Long, Long>()
        var currentType = 0
        val typeList = listOf(501, 502, 503)
        val startDate = 1614097999402
        val endDate = 1614761199000
        val dateList = makeDateInterval(IntervalDto(startDate, endDate))
        val xlWB = XSSFWorkbook()
        val boldCellStyle: XSSFCellStyle = xlWB.createCellStyle()
        val isBold: XSSFFont = xlWB.createFont()
        isBold.bold = true
        boldCellStyle.setFont(isBold)
        val pattern = SimpleDateFormat("dd-MM-yyyy")
        typeList.forEach { type ->
            if (currentType != type) {
                totalInRegions = mutableMapOf()
            }
            val xlWS = when(type) {
                501 -> xlWB.createSheet("Fuqoro")
                503 -> xlWB.createSheet("IR")
                else -> xlWB.createSheet("LBG")
            }
            var totalInDate = 0
            dateList.forEachIndexed { i, date ->
                var endedDate: Long? = null
                endedDate = if (dateList.size > 1 && dateList.lastIndex != i) {
                    dateList[i] + DateUtil.DAY_MILLISECONDS - 1
                } else {
                    endDate
                }
                val startedDate = if (i == 0) {
                    startDate
                } else {
                    dateList[i]
                }
                println("$startedDate, $endedDate")
                val list = getItems(date, type.toLong())
                list.forEach {
                    if (totalInRegions[it.id] == null) {
                        totalInRegions[it.id] = it.total
                    } else {
                        val oldVal = totalInRegions[it.id]
                        totalInRegions[it.id] = it.total + oldVal!!
                    }
                }
                if (i == 0) {
                    for (j in 0..list.size+5) {
                        xlWS.createRow(j)
                    }
                }
                xlWS.createFreezePane(0, 1)
                list.forEachIndexed { index, dto ->
                    if (index == 0) {
                        val headerNameCell = xlWS.getRow(0).createCell(0)
                        headerNameCell.setCellValue("Xudud")
                        xlWS.autoSizeColumn(0)
                        headerNameCell.cellStyle = boldCellStyle
                        val headerDateCell = xlWS.getRow(index).createCell(i+1)
                        headerDateCell.setCellValue(pattern.format(date))
                        headerDateCell.cellStyle = boldCellStyle
                    }
                    xlWS.getRow(index+1).createCell(0).setCellValue(dto.name)
                    xlWS.getRow(index+1).createCell(i+1).setCellValue(dto.total.toString())
                    xlWS.autoSizeColumn(i+1)
                    totalInDate += dto.total.toInt()
                }
                val footerTotalCell = xlWS.getRow(list.lastIndex+2).createCell(i+1)
                footerTotalCell.setCellValue(totalInDate.toString())
                footerTotalCell.cellStyle = boldCellStyle
                totalInDate = 0
                val headerTotalCell = xlWS.getRow(0).createCell(dateList.lastIndex+2)
                headerTotalCell.setCellValue("Jami")
                headerTotalCell.cellStyle = boldCellStyle

                val footerTotalNamedCell = xlWS.getRow(list.lastIndex+2).createCell(0)
                footerTotalNamedCell.setCellValue("Jami")
                footerTotalNamedCell.cellStyle = boldCellStyle
                currentType = type
                for ((k, v) in totalInRegions.values.toList().withIndex()) {
                    xlWS.getRow(k+1).createCell(dateList.lastIndex+2).setCellValue(v.toString())
                }
            }
        }
        val byteArrayOS = ByteArrayOutputStream()
        xlWB.write(byteArrayOS)
        xlWB.close()


        val headers = HttpHeaders().apply {
            contentType = MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8")
            contentDisposition = ContentDisposition.builder("inline").filename("${LocalDate.now()}.xlsx").build()
            contentLength = byteArrayOS.size().toLong()
        }

        return ResponseEntity(byteArrayOS.toByteArray(), headers, HttpStatus.OK)
    }
}

fun makeDateInterval(dto: IntervalDto): List<Long> {
    var startDate = Instant.ofEpochMilli(dto.startDate).atZone(ZoneId.systemDefault()).toLocalDate().atStartOfDay().toEpochSecond(
        ZoneOffset.UTC).times(1000)
    val endDate = Instant.ofEpochMilli(dto.endDate).atZone(ZoneId.systemDefault()).toLocalDate().atStartOfDay().toEpochSecond(
        ZoneOffset.UTC).times(1000)
    val dateList = mutableListOf<Long>()
    while (startDate <= endDate) {
        dateList.add(startDate)
        startDate += DateUtil.DAY_MILLISECONDS
    }
    return dateList
}