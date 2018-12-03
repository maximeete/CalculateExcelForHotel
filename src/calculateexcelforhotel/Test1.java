/*
这个程序的主要慢，就慢在dealExcel这个方法那里
 */
package calculateexcelforhotel;

import java.awt.BorderLayout;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetAdapter;
import java.awt.dnd.DropTargetDropEvent;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.*;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;

/**
 *
 * @author 91152
 */
public class Test1 {

    private JFrame jf = null;
    private JTextArea jt = new JTextArea();

    private void init() {
        jf = new JFrame("自动计算房间");
        ImageIcon im = new ImageIcon("22.png");
        JLabel jl = new JLabel(im);
        
        jf.setSize(400, 400);
        jt.setBackground(new java.awt.Color(220,220,220));
        jt.setLineWrap(true);
        jt.setText("\n第一步：在门卡系统里导出excel文件"
                + "\n第二步：将excel文件打开，然后另存一下（这个主要是因为导出来的文件"
                + "天生有问题，这个门卡软件的问题）"
                + "\n第三步：将另存的文件拖入到本软件的界面内，输入需要查询的月份"
                + "\n第四步：提示成功后，在当前文件夹下打开ben.xls查看详情");
        jt.setEditable(false);
        jf.add(jt);
        jf.add(jl,BorderLayout.SOUTH);
        
        jf.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        jf.setLocationRelativeTo(null);
        jf.setVisible(true);
        //新建一个DropTarget对象
        new DropTarget(jf, DnDConstants.ACTION_COPY, new BenDropTargetListener());
    }

    public static void main(String[] args) {
        new Test1().init();
    }

    //处理Excel
    private void dealExcel(File file,int num,String houZhui) {
        int fangJianNum = 0; //正常的房间间数
        int fangJianNum1 = 0; //低于30分钟的间数
        int fangJianNum2 = 0; //24到25之间小时的房间数
        Row row;
        Cell cell1, cell2;
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm");
        String ss1;
        String ss2;
        Date date1 = null;//入住日期
        Date date2 = null;//退房日期
        long jianGeShiJian;
        Workbook wb = null;
        CellStyle cellstyle;
        CellStyle cellstyle2;
        CellStyle cellstyle3;
        Sheet sheet;
        int lastRow;
        Calendar calendar1 = Calendar.getInstance();//用于分析date1，登记的日期
        Calendar calendar2 = Calendar.getInstance();//用于分析date1，退房日期的
        Calendar calendar3 = Calendar.getInstance();//使用软件当天的日期，不要进行任何的set改变
        Calendar calendar4 = Calendar.getInstance();//用于分析月底的天数，因为设计到setDate，所以，才用到这个变量
        int a1;//登记日期的月份
        int a2;//退房日期的月份
        int a3;//查询日期的月份
        int tianShu;//一个在计算总天数的时候用于过渡的天数
        Row rowL;//用于在Excel底部写分析结果
        String luJing;//导出文件时的路径
        double ll2;//用于处理情况2-3和2-4的double值
        String houZhui1 = houZhui;//文件的后缀名

        try {
            wb = WorkbookFactory.create(file);
        } catch (IOException ex) {
            Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (EncryptedDocumentException ex) {
            Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        cellstyle = wb.createCellStyle();
        cellstyle2 = wb.createCellStyle();
        cellstyle3 = wb.createCellStyle();
        cellstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellstyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellstyle3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellstyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellstyle2.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        cellstyle3.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        sheet = wb.getSheetAt(0);
        lastRow = sheet.getLastRowNum();

        //关键就是这个for循环的处理，其他的都是筹备！！！！！！！！！----------------------------
        for (int i = 0; i < lastRow; i++) {
            //开始处理单元格
            row = sheet.getRow(i + 1);
            cell1 = row.getCell(8);
            cell2 = row.getCell(12);
            
            ss1 = cell1.getStringCellValue();
             try {
                    date1 = sdf.parse(ss1);
                } catch (ParseException ex) {
                    Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
                }
             calendar1.setTime(date1);
             a1 = calendar1.get(Calendar.MONTH);
             a3 = calendar3.get(Calendar.MONTH);
             
            //情况4
            if(cell2 == null && (a1+1) == num && (a3+1) == num){
                tianShu = calendar3.get(Calendar.DATE)-calendar1.get(Calendar.DATE);
                fangJianNum = fangJianNum + tianShu;
                row.createCell(18).setCellValue(tianShu);
                continue;}
            //情况6
            else if(cell2 == null && (a1+1) == num && (a3+1) != num){
                calendar4.set(Calendar.MONTH, a1);//放入入住那天的月份
                calendar4.set(Calendar.DATE, 1);//把日期设置为当月第一天  
                calendar4.roll(Calendar.DATE, -1);//日期回滚一天，也就是最后一天  
                tianShu = calendar4.get(Calendar.DATE) - calendar1.get(Calendar.DATE)+1;
                fangJianNum = fangJianNum + tianShu;
                row.createCell(18).setCellValue(tianShu);
                continue;
            }
            //情况5
            else if(cell2 == null && (num-1) == (a1+1)){
                tianShu = calendar3.get(Calendar.DATE) -1;
                fangJianNum = fangJianNum + tianShu;
                row.createCell(18).setCellValue(tianShu);
                continue;
                    }
            else if(cell2 == null){
                for(int ii = 0;ii<12;ii++){
                        //将日期为空，又不属于情况4和情况5的行弄成橘色，也就是说这些是不属于任何情况的
                        row.getCell(ii).setCellStyle(cellstyle3);
                    }
                for(int ii = 13;ii<row.getLastCellNum();ii++){
                        //将日期为空，又不属于情况4和情况5的行弄成橘色，也就是说这些是不属于任何情况的
                        row.getCell(ii).setCellStyle(cellstyle3);
                    }
                continue;
            }
            
            //cell2不为空才往下计算
            ss2 = cell2.getStringCellValue();
            try {
                date2 = sdf.parse(ss2);
            } catch (ParseException ex) {
                Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
            }
            
            calendar2.setTime(date2);
            a2 = calendar2.get(Calendar.MONTH);
            //情况1
            if((a2+1) == num && (a1+1) == (num -1)){
                tianShu = calendar2.get(Calendar.DATE)-1;
                fangJianNum = fangJianNum + tianShu;
                row.createCell(18).setCellValue(tianShu);
            }
            //情况2
            else if((a2+1) == num && (a1+1) == (num)){
                jianGeShiJian = date2.getTime() - date1.getTime();
                //30分钟就是1,800,000哈，一天是24*60*60*1000=86,400,000
                //情况2-1
                if(jianGeShiJian<=1800000){
                    fangJianNum1++;
                    for(int ii = 0;ii<row.getLastCellNum();ii++){
                        //所有低于30分钟的都给他们弄成淡绿色！
                        row.getCell(ii).setCellStyle(cellstyle);
                    }
                    row.createCell(18).setCellValue(0);
                }
                //情况2-2
                else if (jianGeShiJian >1800000 && jianGeShiJian<=86400000){
                    fangJianNum++;
                    row.createCell(18).setCellValue(1);
                }
                //最后一天超过24小时，小于25小时，一小时为3，600,000，这种情况只算一天哦
                //情况2-3
                else if (jianGeShiJian > 86400000 && (jianGeShiJian%86400000)<=3600000) {
                    ll2 = (double) jianGeShiJian / (1000 * 60 * 60 * 24);
                    tianShu = (int) Math.ceil(ll2)-1;
                    fangJianNum = fangJianNum + tianShu;
                    fangJianNum2++;
                    for(int ii = 0;ii<row.getLastCellNum();ii++){
                        //所有最后一天超过在1小时内的都给他们弄成淡黄色！
                        row.getCell(ii).setCellStyle(cellstyle2);
                    }
                    row.createCell(18).setCellValue(tianShu);
                } 
                //最后一天超过25个小时的
                //情况2-4
                else{
                    ll2 = (double) jianGeShiJian / (1000 * 60 * 60 * 24);
                    tianShu = (int) Math.ceil(ll2);
                    fangJianNum = fangJianNum + tianShu;
                    row.createCell(18).setCellValue(tianShu);
                        } 
                    }
            //情况3
            else if((a2+1) == (num+1) && (a1+1) == (num)){
                calendar4.set(Calendar.MONTH, num-1);
                calendar4.set(Calendar.DATE, 1);//把日期设置为当月第一天  
                calendar4.roll(Calendar.DATE, -1);//日期回滚一天，也就是最后一天  
                tianShu = calendar4.get(Calendar.DATE) - calendar1.get(Calendar.DATE)+1;
                fangJianNum = fangJianNum + tianShu;
                row.createCell(18).setCellValue(tianShu);
                    }
            //所有最后一天超过在1小时内的都给他们弄成橘色！
            else{
                for(int ii = 0;ii<row.getLastCellNum();ii++){
                        row.getCell(ii).setCellStyle(cellstyle3);
                    }
            }
            }
        //---------------------------------------------------------------------
        
        rowL = sheet.createRow(lastRow +2);
        rowL.createCell(0).setCellValue(num +"月份总间数：" + fangJianNum + "间");
        rowL = sheet.createRow(lastRow +3);
        rowL.createCell(0).setCellValue("30分钟内的总间数：" + fangJianNum1 + "间，表格中标了淡绿色，并不算间数");
        rowL = sheet.createRow(lastRow +4);
        rowL.createCell(0).setCellValue("最后一天超了1小时内的为：" + fangJianNum2 + "间，表格中标了淡黄色，这1小时不算间数");
        rowL = sheet.createRow(lastRow +6);
        rowL.createCell(0).setCellValue("橙色的表示不在查询的月份范围");
        rowL = sheet.createRow(lastRow +7);
        rowL.createCell(0).setCellValue("表格最右边已经列出每一行记录所对应的房间数，用于进行比对");
        
        
        //导出文件
        luJing = file.getParent()+"\\ben." + houZhui1;
        try {
            try (FileOutputStream fileout = new FileOutputStream(luJing)) {
                wb.write(fileout);
            }
            wb.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    class BenDropTargetListener extends DropTargetAdapter {

        @Override
        public void drop(DropTargetDropEvent dtde) {
            File file;
            Transferable tf;
            DataFlavor[] df;
            DataFlavor ddf;
            List fileList;
            String filename;//文件名
            String houZhui;//后缀名
            int a;//用于过渡
            String yueFen;//用于得到用户输入的值，String
            int ben1;//将yueFen转化为int值

            //下面这一句不写就会报错，导致下面无法获取fileList
            dtde.acceptDrop(DnDConstants.ACTION_COPY);

            tf = dtde.getTransferable();
            df = dtde.getCurrentDataFlavors();
            //为什么要做下面的判断？因为非常有可能拖入其他的类型，比如文本类型
            //那么，如果没有javaFileListeFlavor类型的话，就没有下面的步骤了
            for (int i = 0; i < df.length; i++) {
                ddf = df[i];
                if (ddf.equals(DataFlavor.javaFileListFlavor)) {
                    try {
                        fileList = (List) tf.getTransferData(ddf);
                        for (Object f : fileList) {
                            file = (File) f;
                            filename = file.getAbsolutePath();
                            a = filename.lastIndexOf(".");
                            houZhui = filename.substring(a + 1);
                            if (!houZhui.equals("xls") && !houZhui.equals("xlsx")) {
                                JOptionPane.showMessageDialog(jf, "你拖入的文件不是excel文件");
                                continue;
                            }
                            yueFen = JOptionPane.showInputDialog(jf, 
                                    "请问需要查哪一个月份的？", "输入月份", JOptionPane.PLAIN_MESSAGE);
                            ben1 = dealString(yueFen);
                            if(ben1 == 1){
                                JOptionPane.showMessageDialog(jf, "你输入的数字有误，请重新拖入文件");
                                dtde.dropComplete(true);
                            }else{
                            
                            dealExcel(file,ben1,houZhui);
                            JOptionPane.showMessageDialog(jf, "已经成功导出，请查看当前文件夹下的ben.xls");
                            }
                        }
                    } catch (UnsupportedFlavorException | IOException ex) {
                        Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
                    }

                }
            }
            dtde.dropComplete(true);
        }

    }
    
    private int dealString(String s){
        int a;
        int[] intArray= {1,2,3,4,5,6,7,8,9,10,11,12};
        
        try{
        a = Integer.parseInt(s);}
        catch(NumberFormatException e){
            e.printStackTrace();
            return 1;
        }
        
        for(int benIntArray:intArray ){
            if(a == benIntArray){
                return a;
            }
        }
        return 1;
    }
}
