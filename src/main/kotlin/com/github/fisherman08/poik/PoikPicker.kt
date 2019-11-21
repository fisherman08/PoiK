package com.github.fisherman08.poik

import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddressBase
import java.io.*

class PoikPicker(input: InputStream, sheetName: String = "") {

    private val workbook: Workbook = WorkbookFactory.create(input)
    private val sheet: Sheet = if (sheetName != "") {
        workbook[sheetName] ?: throw SheetNotFoundException(sheetName)
    } else {
        workbook.first()
    }


    constructor(file: File, sheetName: String = ""): this(FileInputStream(file), sheetName)

    constructor(bytes: ByteArray, sheetName: String = ""): this(ByteArrayInputStream(bytes), sheetName)

    val maxRow: Int = sheet.lastRowNum
    val maxColumn: Int

    init {
        // 最初にこのシートに存在する最大の列数を取得しておく
        var result = 0
        for (rowNo in sheet.firstRowNum..sheet.lastRowNum) {
            val maxColumnNo = sheet[rowNo]?.lastCellNum?.toInt() ?: result
            if (maxColumnNo > result) {
                result = maxColumnNo
            }

        }

        maxColumn = result
    }

    infix fun exec(action: (picker: PoikPicker)-> Unit){
        action(this)
    }


    /**
     * ファイルへのIOを閉じる
     *
     */
    fun close() {
        workbook.close()
    }

    /**
     * 行と列指定で値をStringで取得する
     */
    fun string(row: Int, column: Int): String {

        try {
            val cell =  sheet[row, column]  ?: return ""

            if (cell.cellTypeEnum == CellType.NUMERIC) {
                // 数値型のセルをそのまま取得すると勝手にいらん小数点がついてしまうので文字列型にしてから取得する
                cell.setCellType(CellType.STRING)
            }

            return cell.stringCellValue

        } catch (e: Exception) {
            throw InvalidValueException(e, sheet.sheetName, row, column)
        }

    }

    /**
     * 行と列指定で値をDoubleで取得する
     */
    fun double(row: Int, column: Int): Double {
        try {

            val cell = sheet[row, column] ?: return 0.0

            return if(cell.cellTypeEnum == CellType.STRING){
                // カラムが文字列
                cell.stringCellValue.toDouble()
            } else {
                cell.numericCellValue
            }

        } catch (e: Exception) {
            throw InvalidValueException(e, sheet.sheetName, row, column)
        }
    }

    /**
     * 行と列指定で値をIntで取得する。少数以下は切り捨てられますよ。
     */
    fun int(row: Int, column: Int): Int {
        return double(row, column).toInt()
    }



    /**
     * 指定されたシートが見つからなかった時の例外
     */
    class SheetNotFoundException(sheetName: String) : Exception("Sheet: '$sheetName' does not exist")

    /**
     * 指定されたセルが見つからなかった時の例外
     */
    class CellNotFoundException : Exception("Cell is Null")

    /**
     * valueをうまく取得できなかった時の例外
     */
    class InvalidValueException(error: Exception, sheetName: String, row: Int, column: Int) : Exception("Cannot get Cell value at position[$row, $column] of sheet: '$sheetName'", error)


// 以下はpoiの使い勝手を良くする拡張関数たち

    /**
     * workbook["シート名"]でアクセスできるようにする
     */
    operator fun Workbook.get(sheetName: String): Sheet? {
        return this.getSheet(sheetName)
    }

    operator fun Sheet.get(row: Int): Row? {
        return getRow(row)
    }

    /**
     * sheet[row, column]でアクセスできるようにする
     */
    operator fun Sheet.get(row: Int, column: Int): Cell? {
        val cell = this.getRow(row)?.getCell(column) ?: return null

        // シート内のmergedRegionの一覧から自分(Cell)が含まれているものを探す
        merged_region_all@ for (mergedRegion in this.mergedRegions) {
            if (!mergedRegion.isInRange(cell)) {
                // このmergedRegionには含まれていない
                continue
            }

            if (!mergedRegion.isFirstCell(cell)) {
                // 結合されたセルだったら本来のセルは無視して結合された中の一番左上のセルの値で上書きする
                val cellToDisplay = this[mergedRegion.firstRow, mergedRegion.firstColumn] ?: cell
                cell.copyValues(cellToDisplay)
            }

            // 同時に2つ以上のmergedRegionに含まれることはないので、マッチした時点でループを止める
            break@merged_region_all
        }

        return cell
    }

    fun Row.getCells(): Array<Cell> {
        val result: MutableList<Cell> = mutableListOf()
        for (i in firstCellNum..lastCellNum) {
            val cell = sheet[rowNum, i] ?: continue
            result.add(cell)
        }
        return result.toTypedArray()
    }


    /**
     * セルタイプに応じてセルの値をコピーする
     */
    fun Cell.copyValues(source: Cell) {

        when (source.cellTypeEnum) {
            CellType.NUMERIC -> setCellValue(source.numericCellValue)
            CellType.STRING -> setCellValue(source.stringCellValue)
            CellType.BOOLEAN -> setCellValue(source.booleanCellValue)
            CellType.FORMULA -> setCellFormula(source.cellFormula)
            else -> {
            }
        }
    }

    /**
     * 結合セル内の最初のセル(一番左上のセル)かどうかを判定
     * 引数で渡されたCellの行と列が結合セルの最初の行と列ならtrue、それ以外ならfalse
     * @param cell Cell
     * @return Boolean
     */
    fun CellRangeAddressBase.isFirstCell(cell: Cell): Boolean {
        return (cell.rowIndex == this.firstRow && cell.columnIndex == this.firstColumn)
    }


}



