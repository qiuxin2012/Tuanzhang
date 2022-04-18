package com.github.qiuxin2012.tuan


import org.apache.poi.ss.usermodel.{Row, Sheet}

import java.io.{File, FileInputStream}
import java.io.FileOutputStream
import java.io.IOException
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import scala.collection.mutable.ArrayBuffer

object Reorder {
  def main(args: Array[String]): Unit = {
    val filePath = args(0)
//    val filePath = "C:\\Users\\xinqiu\\Downloads\\订单20_50_50.xlsx.xlsx"
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
    var out: FileOutputStream = null
    try { // 获取总列数
      val outputPath = args(1)
//      val outputPath = new File("C:\\Users\\xinqiu\\Downloads\\out.xlsx")
      val workBook = new XSSFWorkbook()
      val partSplit = Array(0, 55, 79, 97, 128)
      (1 until partSplit.length).foreach{part =>
        writeToSheet(orders.filter(o => o.no >= (partSplit(part-1)+1) && o.no <= partSplit(part)),
          workBook.createSheet(s"${partSplit(part-1)+1}-${partSplit(part)}号"))
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

  def writeToSheet(orders: Array[Order], sheet: Sheet): Unit = {

    var noItems = scala.collection.mutable.Map(orders.map(_.item).distinct.map(i => (i, 0)): _*)
    val partItems = noItems.clone()
    var currentNo = orders(0).no
    val title = Array("名字","商品","数量","电话","楼栋","房号")
    set(sheet.createRow(0), title)
    var i = 0
    var offset = 1
    while(i < orders.length) {
      val order = orders(i)
      noItems(order.item) = noItems(order.item) + order.num
      partItems(order.item) = partItems(order.item) + order.num
      val row = sheet.createRow(i + offset)
      set(row, order)
      if (i+1 == orders.length || order.no != orders(i+1).no) {
        offset += 1
        set(sheet.createRow(i + offset), Array(s"${order.no}号")
          ++ noItems.flatMap(v => Array(s"${v._1}", s"${v._2}份")))
        noItems = scala.collection.mutable.Map(orders.map(_.item).distinct.map(i => (i, 0)): _*)
        currentNo = orders(i).no
      }
      i += 1
    }
    set(sheet.createRow(i + offset + 1), Array(sheet.getSheetName)
      ++ partItems.flatMap(v => Array(s"${v._1}", s"${v._2}份")))
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

  def set(r: Row, a: Array[String]): Unit = {
//    println(r.getRowNum)
    a.indices.foreach{l =>
      r.createCell(l).setCellValue(a(l))
    }
  }
}

case class Order(name: String, item: String, num:Int, tel: String, no: Int, room: String)