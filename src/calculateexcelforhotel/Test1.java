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
        jf.setSize(400, 400);
        
        ImageIcon im = new ImageIcon("22.png");
        JLabel jl = new JLabel(im);
        
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
    private void dealExcel(File file,int num) {
        int fangJianNum = 0; //正常的房间间数
        int fangJianNum1 = 0; //低于30分钟的间数
        int fangJianNum2 = 0; //24到25之间小时的房间数
        //int dbg = 0; //用于debug，不用删除
        Row row = null;
        Cell cell1, cell2;
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm");
        String ss1 = null;
        String ss2 = null;
        Date date1 = null;
        Date date2 = null;
        long jianGeShiJian;
        Workbook wb = null;

        try {
            wb = WorkbookFactory.create(file);
        } catch (IOException ex) {
            Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (EncryptedDocumentException ex) {
            Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        CellStyle cellstyle = wb.createCellStyle();
        CellStyle cellstyle2 = wb.createCellStyle();
        CellStyle cellstyle3 = wb.createCellStyle();
        cellstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellstyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellstyle3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellstyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellstyle2.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        cellstyle3.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        Sheet sheet = wb.getSheetAt(0);
        int lastRow = sheet.getLastRowNum();

        //关键就是这个for循环的处理，其他的都是筹备！！！！！！！！！----------------------------
        for (int i = 0; i < lastRow; i++) {
            //开始处理单元格
            row = sheet.getRow(i + 1);
            cell1 = row.getCell(8);
            cell2 = row.getCell(12);
            
            ss1 = cell1.getStringCellValue();
             try {
                    date1 = sdf.parse(ss1);//入住日期
                } catch (ParseException ex) {
                    Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
                }
             Calendar calendar1 = Calendar.getInstance();
             calendar1.setTime(date1);
             int a1 = calendar1.get(Calendar.MONTH);
             
            //情况4
            if(cell2 == null && (a1+1) == num){
               
                Calendar calendar3 = Calendar.getInstance();
                int tianShu = calendar3.get(Calendar.DATE)-calendar1.get(Calendar.DATE);
                fangJianNum = fangJianNum + tianShu;
                row.createCell(18).setCellValue(tianShu);
                continue;}
            //情况5
            else if(cell2 == null && (num-1) == (a1+1)){
                Calendar calendar3 = Calendar.getInstance();
                int tianShu = calendar3.get(Calendar.DATE) -1;
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
            
            
            ss2 = cell2.getStringCellValue();
            try {
                date2 = sdf.parse(ss2);//退房日期
            } catch (ParseException ex) {
                Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
            }
            Calendar calendar2 = Calendar.getInstance();
            calendar2.setTime(date2);
            int a2 = calendar2.get(Calendar.MONTH);
            //情况1
            if((a2+1) == num && (a1+1) == (num -1)){
                int tianShu = calendar2.get(Calendar.DATE)-1;
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
                    double ll2 = (double) jianGeShiJian / (1000 * 60 * 60 * 24);
                    int tianShu = (int) Math.ceil(ll2)-1;
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
                    double ll2 = (double) jianGeShiJian / (1000 * 60 * 60 * 24);
                    int tianShu = (int) Math.ceil(ll2);
                    fangJianNum = fangJianNum + tianShu;
                    row.createCell(18).setCellValue(tianShu);
                        } 
                    }
            //情况3
            else if((a2+1) == (num+1) && (a1+1) == (num)){
                Calendar a = Calendar.getInstance();  
                a.set(Calendar.MONTH, num-1);
                a.set(Calendar.DATE, 1);//把日期设置为当月第一天  
                a.roll(Calendar.DATE, -1);//日期回滚一天，也就是最后一天  
                int maxDate = a.get(Calendar.DATE);
                int tianShu = maxDate - calendar1.get(Calendar.DATE)+1;
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
        
        Row rowL = sheet.createRow(lastRow +2);
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
        String luJing = file.getParent()+"\\ben.xls";
        try {
            FileOutputStream fileout = new FileOutputStream(luJing);
            wb.write(fileout);
            fileout.close();
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
            File file = null;

            //下面这一句不写就会报错，导致下面无法获取fileList
            dtde.acceptDrop(DnDConstants.ACTION_COPY);

            Transferable tf = dtde.getTransferable();
            DataFlavor[] df = dtde.getCurrentDataFlavors();
            //为什么要做下面的判断？因为非常有可能拖入其他的类型，比如文本类型
            //那么，如果没有javaFileListeFlavor类型的话，就没有下面的步骤了
            for (int i = 0; i < df.length; i++) {
                DataFlavor ddf = df[i];
                if (ddf.equals(DataFlavor.javaFileListFlavor)) {
                    try {
                        List fileList = (List) tf.getTransferData(ddf);
                        for (Object f : fileList) {
                            file = (File) f;
                            String filename = file.getAbsolutePath();
                            int a = filename.lastIndexOf(".");
                            String houZhui = filename.substring(a + 1);
                            if (!houZhui.equals("xls")) {
                                JOptionPane.showMessageDialog(jf, "你拖入的文件不是xls格式的");
                                continue;
                            }
                            String yueFen = JOptionPane.showInputDialog(jf, 
                                    "请问需要查哪一个月份的？", "输入月份", JOptionPane.PLAIN_MESSAGE);
                            int ben1 = dealString(yueFen);
                            if(ben1 == 1){
                                JOptionPane.showMessageDialog(jf, "你输入的数字有误，请重新拖入文件");
                                dtde.dropComplete(true);
                            }else{
                            
                            dealExcel(file,ben1);
                            JOptionPane.showMessageDialog(jf, "已经成功导出，请查看当前文件夹下的ben.xls");
                            }
                        }
                    } catch (UnsupportedFlavorException ex) {
                        Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(Test1.class.getName()).log(Level.SEVERE, null, ex);
                    }

                }
            }
            dtde.dropComplete(true);
        }

    }
    
    private int dealString(String s){
        int a;
        try{
        a = Integer.parseInt(s);}
        catch(Exception e){
            e.printStackTrace();
            return 1;
        }
        switch (a){
            case 1:
            case 2:
            case 3:
            case 4:
            case 5:
            case 6:
            case 7:
            case 8:
            case 9:
            case 10:
            case 11:
            case 12:
                return a;
            default:
                return 1;
                
        }
    }
}
