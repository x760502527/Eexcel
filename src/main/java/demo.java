import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by hasee on 2017/12/5.
 */
public class demo {
    public static void  main(String []arg){
        String name= new SimpleDateFormat("YYYY年MM月dd日_hh时mm分ss秒").format(new Date());
        String path = "D:/EExcel" + "\\" + name + ".xls";
        FileOutputStream fileOut = null;
        Workbook workbook = null;


        Eexcel p = new Eexcel();
        p.append("大撒旦撒").append("1",2,2).
                append("dsasd",3,3)
                .next().
                append("dsa",2,2,p.getCellStyleByTable())
                .append("dsadas",5,5,p.getCellStyleByTable());
        p.setCellStyle(p.getCellStyleByTable(),3,0);
        p.setRegionStyle(p.getCellStyleByTable(),0,10,0,10);
        if(p.writeFile(path)){
            System.out.println("生成文件至:"+path);
        }
    }


}
