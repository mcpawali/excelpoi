package com.cpawali.com.r

import java.io.FileOutputStream

import org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType
import org.apache.poi.ss.util.CellRangeAddressList
import org.apache.poi.xssf.usermodel.{XSSFDataValidationHelper, XSSFRow, XSSFWorkbook}

/**
  * Created by chandrashekhar on 2019-10-14
  */

case class TripExcelEntityCommand(columnIndex: Int, columnName: String, isMandatory: Boolean,
                                  isDate: Boolean, isTime: Boolean,
                                  columnEntries: Seq[String] = Nil, headerMessage: String = "")

object TripExcelFields {
  val DATE        = "DATE"
  val SHIFT       = "SHIFT"
  val WEIGHBRIDGE = "WEIGHBRIDGE"
  val STARTTIME   = "STARTTIME"
  val ENDTIME     = "ENDTIME"
}

object ExcelCreation extends App {


  val row_Start = 1
  val row_End   = 1000

  val wb = new XSSFWorkbook()

  val sheet = wb.createSheet()

  val dataValidationHelper = new XSSFDataValidationHelper(sheet)


  def addColumnDescription(tripExcelEntityCommands: List[TripExcelEntityCommand], headerRow: XSSFRow) = {
    val cellStyle = wb.createCellStyle()
    cellStyle.setWrapText(true)
    tripExcelEntityCommands.foreach{ x =>
      sheet.autoSizeColumn(x.columnIndex)
      val cell = headerRow.createCell(x.columnIndex)
      cell.setCellValue(x.headerMessage)
      cell.setCellStyle(cellStyle)
      headerRow.setRowStyle(cellStyle)
    }
  }

  def tripExcelEntityCommands = List(
    TripExcelEntityCommand(0, TripExcelFields.DATE, true, true, false, Nil, "Accepted Date format is yyyy-MM-dd"),
    TripExcelEntityCommand(1, TripExcelFields.SHIFT, true, false, false, List("A", "B")),
    TripExcelEntityCommand(10, TripExcelFields.WEIGHBRIDGE, false, false, false, List("Wb1", "Wb2")),
    TripExcelEntityCommand(11, TripExcelFields.STARTTIME, false, false, true, Nil, "Accepted Time format is hh:mm:SS"),
    TripExcelEntityCommand(12, TripExcelFields.ENDTIME, false, false, true, Nil, "Accepted Time format is hh:mm:SS")
  )

  private def addColumnValidation() = {
    val descriptionRow: XSSFRow = sheet.createRow(row_Start - 1)
    val row: XSSFRow = sheet.createRow(row_Start)

    val headers = addColumnDescription(tripExcelEntityCommands, descriptionRow)

    val file_Name = "Cpawali_Excel.xlsx"
    val fileOut = new FileOutputStream(s"C:\\Cpawali\\Excel\\$file_Name")

    tripExcelEntityCommands.foreach{ x =>
      val cell = row.createCell(x.columnIndex)
      cell.setCellValue(x.columnName)
      sheet.setDefaultColumnWidth(12)


      if (x.columnEntries.nonEmpty) {
        val dropDownEntries = x.columnEntries
        val hidden = wb.createSheet(x.columnName)
        val hiddenSheetIndex = wb.getSheetIndex(hidden)
        dropDownEntries.map{ data =>
          val r = hidden.createRow(dropDownEntries.indexOf(data))
          val c = r.createCell(0)
          c.setCellValue(data)
        }
        val namedCell = wb.createName()
        namedCell.setNameName(x.columnName)
        val formula = x.columnName + "!$A$1:$A$" + dropDownEntries.length
        namedCell.setRefersToFormula(formula)
        val addressList = new CellRangeAddressList(row_Start + 1, row_End, x.columnIndex, x.columnIndex)
        val dvConstraint = dataValidationHelper.createFormulaListConstraint(x.columnName)
        val validation = dataValidationHelper.createValidation(dvConstraint, addressList)

        wb.setSheetHidden(hiddenSheetIndex, true)
        sheet.addValidationData(validation)
      } else if (x.isDate) {
        //row_Start+1 is just to avoid dropdown values on headers
        val addressList = new CellRangeAddressList(row_Start + 1, row_End, x.columnIndex, x.columnIndex)
        val dvConstraint = dataValidationHelper.createDateConstraint(OperatorType.BETWEEN, "DATE(1990,1,1)", "DATE(9999,1,1)", "yyyy-MM-dd")
        val validation = dataValidationHelper.createValidation(dvConstraint, addressList)
        validation.setShowErrorBox(true)
        validation.setEmptyCellAllowed(x.isMandatory)
        validation.setSuppressDropDownArrow(true)
        sheet.addValidationData(validation)
      } else if (x.isTime) {
        //row_Start+1 is just to avoid dropdown values on headers
        val addressList = new CellRangeAddressList(row_Start + 1, row_End, x.columnIndex, x.columnIndex)
        val dvConstraint = dataValidationHelper.createTimeConstraint(OperatorType.BETWEEN, "=TIME(00,00,00)", "=TIME(23,59,59)")
        val validation = dataValidationHelper.createValidation(dvConstraint, addressList)
        validation.setShowErrorBox(true)
        validation.setEmptyCellAllowed(x.isMandatory)
        validation.setSuppressDropDownArrow(true)
        sheet.addValidationData(validation)
      } else {

        //row_Start+1 is just to avoid dropdown values on headers
        val addressList = new CellRangeAddressList(row_Start + 1, row_End, x.columnIndex, x.columnIndex)
        val dvConstraint = dataValidationHelper.createTextLengthConstraint(OperatorType.GREATER_THAN, "0", "100")
        val validation = dataValidationHelper.createValidation(dvConstraint, addressList)
        validation.setShowErrorBox(true)
        validation.setEmptyCellAllowed(x.isMandatory)
        validation.setSuppressDropDownArrow(true)
        sheet.addValidationData(validation)
      }
    }
    wb.write(fileOut)
    _ = fileOut.close()
  }

  addColumnValidation()
  Thread.sleep(2000)
}