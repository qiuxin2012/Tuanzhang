import org.apache.poi.ss.usermodel.{BorderStyle, Row, Sheet}

import java.io.{File, FileInputStream}
import java.io.FileOutputStream
import java.io.IOException
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFWorkbook}

import java.awt.Color
import scala.collection.mutable.ArrayBuffer

object Reorder {
  def main(args: Array[String]): Unit = {
//    val filePath = args(0)
    val filePath = "C:\\Users\\xinqiu\\Downloads\\订单15_28_40.xlsx"
    val fs = new FileInputStream(filePath)
    val xssfWorkbook: XSSFWorkbook = new XSSFWorkbook(fs)
    val order = ArrayBuffer[Order]()
    //获取表格第一个sheet
    val xssfSheet = xssfWorkbook.getSheetAt(0)
    val titleRow = xssfSheet.getRow(0)
    for(row <- 1 to xssfSheet.getLastRowNum()){
      //获取表格每一行
      val xssfRow = xssfSheet.getRow(row)
      val name = xssfRow.getCell(0).toString.trim
      print(name)
      val item = xssfRow.getCell(4).toString.trim
      val num = xssfRow.getCell(6).toString.trim.toDouble.toInt
      val tel = xssfRow.getCell(18).toString.trim
      val noO = xssfRow.getCell(22).toString.trim
      var no = ""
      var i = 0
      while(i < noO.length){
        if (noO.charAt(i) >= 48 && noO.charAt(i) <= 57) {
          no += noO.charAt(i)
          i += 1
        } else {
          i = Int.MaxValue
        }
      }
      val room = xssfRow.getCell(21).toString.trim
      println(name)
      order.append(Order(name, item,
        num, tel,
        no.toInt, room))
    }
    val orders = order.toArray.sortBy(_.room).sortBy(_.no)
    val names = orders.map(_.name).distinct
    println(names.mkString("\n"))
      //.filter(v => v.item == "悦鲜活牛奶+简醇酸奶套装(牛奶950ml*6瓶/酸奶150克*16袋)" || v.item == "白小纯常温奶(180克*16袋)")
    var out: FileOutputStream = null
    try { // 获取总列数
//      val outputPath = args(1)
      val outputPath = new File("C:\\Users\\xinqiu\\Downloads\\out.xlsx")
      val workBook = new XSSFWorkbook()
      val cellStyle = workBook.createCellStyle()
      //val partSplit = Array(0, 55, 79, 97, 128)
      val partSplit = Array(0, 128)
//      val partSplit = Array(0, 59, 128)
      (1 until partSplit.length).foreach{part =>
        writeToSheet(orders.filter(o => o.no >= (partSplit(part-1)+1) && o.no <= partSplit(part)),
          workBook.createSheet(s"${partSplit(part-1)+1}-${partSplit(part)}号"), cellStyle)
      }

      // 创建文件输出流，输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
      out = new FileOutputStream(outputPath)
      workBook.write(out)
    } catch {
      case e: Exception =>
        e.printStackTrace()
    } finally try if (out != null) {
      out.flush
      out.close
    }
    catch {
      case e: IOException =>
        e.printStackTrace()
    }
  }

  def writeToSheet(orders: Array[Order], sheet: Sheet, cellStyle: XSSFCellStyle): Unit = {

    var noItems = scala.collection.mutable.Map(orders.map(_.item).distinct.map(i => (i, 0)): _*)
    val partItems = noItems.clone()
    var currentNo = orders(0).no
    val title = Array("名字","商品","数量","楼栋","房号","电话")
    set(sheet.createRow(0), title)
    var i = 0
    var offset = 1
    import org.apache.poi.xssf.usermodel.XSSFCellStyle
    var orderName = orders(0).name
    while(i < orders.length) {
      val order = orders(i)
      noItems(order.item) = noItems(order.item) + order.num
      partItems(order.item) = partItems(order.item) + order.num

      val row = sheet.createRow(i + offset)
      if (orderName != order.name) {
        cellStyle.setBorderBottom(BorderStyle.HAIR)
        setWithBorder(row, order, cellStyle)
      } else {
        cellStyle.setBorderBottom(BorderStyle.NONE)
        setWithBorder(row, order, cellStyle)
      }
      if (i+1 == orders.length || order.no != orders(i+1).no) {
        offset += 1
        // set(sheet.createRow(i + offset), Array(s"${order.no}号")
        //   ++ noItems.filter(_._2 != 0).flatMap(v => Array(s"${v._1}", s"${v._2}份")))
//        set(sheet.createRow(i + offset), Array(s"${order.no}号")
//          ++ noItems.filter(_._2 != 0).flatMap(v => Array(s"${v._1} ${v._2}份")))
//        offset += 1
        noItems = scala.collection.mutable.Map(orders.map(_.item).distinct.map(i => (i, 0)): _*)
        currentNo = orders(i).no
      }
      i += 1
    }
    // set(sheet.createRow(i + offset + 1), Array("总"+sheet.getSheetName)
    //   ++ partItems.flatMap(v => Array(s"${v._1}", s"${v._2}份")))
    set(sheet.createRow(i + offset + 1), Array("总"+sheet.getSheetName)
      ++ partItems.flatMap(v => Array(s"${v._1} ${v._2}份")))
  }

  def set(r: Row, o: Order): Unit = {
//    println(r.getRowNum)
    r.createCell(0).setCellValue(o.name)
    r.createCell(1).setCellValue(o.item)
    r.createCell(2).setCellValue(o.num)
    r.createCell(3).setCellValue(o.tel)
    r.createCell(4).setCellValue(o.no)
    r.createCell(5).setCellValue(o.room)
  }

  def setWithBorder(r: Row, o: Order, cellStyle: XSSFCellStyle): Unit = {
    //    println(r.getRowNum)
    val r0 = r.createCell(0)
    r0.setCellStyle(cellStyle)
    r0.setCellValue(o.name)
    val r1 = r.createCell(1)
    r1.setCellValue(o.item)
    r1.setCellStyle(cellStyle)
    val r2 = r.createCell(2)
    r2.setCellValue(o.num)
    r2.setCellStyle(cellStyle)
    val r3 = r.createCell(3)
    r3.setCellValue(o.no)
    r3.setCellStyle(cellStyle)
    val r4 = r.createCell(4)
    r4.setCellValue(o.room)
    r4.setCellStyle(cellStyle)
    val r5 = r.createCell(5)
    r5.setCellValue(o.tel)
    r5.setCellStyle(cellStyle)
  }

  def set(r: Row, a: Array[String]): Unit = {
//    println(r.getRowNum)
    a.indices.foreach{l =>
      r.createCell(l).setCellValue(a(l))
    }
  }
}

case class Order(name: String, item: String, num:Int, tel: String, no: Int, room: String)
