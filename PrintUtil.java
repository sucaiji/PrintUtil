package com.sucaiji.util;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;

import javax.imageio.ImageIO;
import javax.print.*;
import javax.print.attribute.DocAttributeSet;
import javax.print.attribute.HashDocAttributeSet;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.standard.OrientationRequested;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class PrintUtil {

    /**
     * 竖屏模式
     */
    public static OrientationRequested PORTRAIT = OrientationRequested.PORTRAIT;

    /**
     * 横屏模式
     */
    public static OrientationRequested LANDSCAPE = OrientationRequested.LANDSCAPE;



    /**
     * 获取全部打印设备信息
     * @return 返回全部能用的打印服务的List
     */
    public static List<PrintService> getDeviceList() {
        // 构建打印请求属性集
        HashPrintRequestAttributeSet pras = new HashPrintRequestAttributeSet();
        // 设置打印格式，因为未确定类型，所以选择autosense
        DocFlavor flavor = DocFlavor.BYTE_ARRAY.AUTOSENSE;
        // 查找所有的可用的打印服务
        PrintService printService[] = PrintServiceLookup.lookupPrintServices(flavor, pras);
        List<PrintService> list = Arrays.asList(printService);
        return list;
    }


    /**
     * 根据文件类型不同调用不同代码去打印
     * @param filePath 文件路径
     */
    public static void print(String filePath) throws Exception {
        PrintService printService = PrintServiceLookup.lookupDefaultPrintService();
        String defaultDeviceName = printService.getName();
        print(filePath, defaultDeviceName);
    }

    /**
     * 额外传入一个 AfterPrint，会在打印完成后调用 afterPrint.run()
     * @param filePath
     * @param afterPrint
     * @throws Exception
     */
    public static void print(String filePath, AfterPrint afterPrint) throws Exception {
        print(filePath);
        afterPrint.run();
    }

    /**
     * 根据文件类型不同调用不同代码去打印
     * @param filePath 文件路径
     * @param deviceName 设备名称，传入哪个设备的名称，就让哪个设备去打印
     */
    public static void print(String filePath, String deviceName) throws Exception{
        List<PrintService> list = getDeviceList();
        PrintService printService = null;
        for (PrintService p : list) {
            if(p.getName().equals(deviceName)) {
                printService = p;
                break;
            }
        }
        if(printService == null) {
            throw new Exception("Device not found");
        }
        String type = filePath.replaceAll(".*\\.","");
        if("jpg".equalsIgnoreCase(type)) {
            normalPrint(new File(filePath), DocFlavor.INPUT_STREAM.JPEG, printService);
            return;
        }
        if("jpeg".equalsIgnoreCase(type)) {
            normalPrint(new File(filePath), DocFlavor.INPUT_STREAM.JPEG, printService);
            return;
        }
        if("gif".equalsIgnoreCase(type)) {
            normalPrint(new File(filePath), DocFlavor.INPUT_STREAM.GIF, printService);
            return;
        }
        if("pdf".equalsIgnoreCase(type)) {
            printPDF(new File(filePath), DocFlavor.INPUT_STREAM.PNG, printService);
            return;
        }
        if("png".equalsIgnoreCase(type)) {
            normalPrint(new File(filePath), DocFlavor.INPUT_STREAM.PNG, printService);
            return;
        }
        if("doc".equalsIgnoreCase(type)) {
            printWord(filePath, deviceName);
            return;
        }
        if("docx".equalsIgnoreCase(type)) {
            printWord(filePath, deviceName);
            return;
        }
        if("xls".equalsIgnoreCase(type)) {
            printExcel(filePath, deviceName);
            return;
        }
        if("xlsx".equalsIgnoreCase(type)) {
            printExcel(filePath, deviceName);
            return;
        }
        if("ppt".equalsIgnoreCase(type)) {
            printPPT(filePath, deviceName);
            return;
        }
        if("pptx".equalsIgnoreCase(type)) {
            printPPT(filePath, deviceName);
            return;
        }

    }

    /**
     * 会在打印完成后调用 afterPrint.run()
     * @param filePath
     * @param deviceName
     * @param afterPrint
     * @throws Exception
     */
    public static void print(String filePath, String deviceName, AfterPrint afterPrint) throws Exception{
        print(filePath, deviceName);
        afterPrint.run();
    }


    /**
     * javase的打印机打印文件，支持jpg,png,gif,pdf等等
     * @param file 要打印的文件
     * @param flavor 打印格式
     */
    private static void normalPrint(File file, DocFlavor flavor) {
        // 定位默认的打印服务
        PrintService service = PrintServiceLookup
                .lookupDefaultPrintService();             // 显示打印对话框
        normalPrint(file, flavor, service);
    }


    private static void normalPrint(File file, DocFlavor flavor, PrintService service) {
        normalPrint(file, flavor, PORTRAIT, service);
    }

    /**
     * javase的打印机打印文件，支持jpg,png,gif等等
     * @param file 要打印的文件
     * @param service 打印机选择
     * @param requested 设定横屏还是竖屏
     * @param flavor 打印格式
     */
    private static void normalPrint(File file, DocFlavor flavor, OrientationRequested requested, PrintService service) {
        // 构建打印请求属性集
        HashPrintRequestAttributeSet pras = new HashPrintRequestAttributeSet();
        pras.add(requested);
        if (service != null) {
            try {
                DocPrintJob job = service.createPrintJob(); // 创建打印作业
                FileInputStream fis = new FileInputStream(file); // 构造待打印的文件流
                DocAttributeSet das = new HashDocAttributeSet();
                Doc doc = new SimpleDoc(fis, flavor, das);
                job.print(doc, pras);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 打印pdf的方法，因为java内置的打印pdf的方法有病，所以首先需要把pdf转换成png，然后打印png
     * @param file 要打印的文件
     * @param flavor 要打印的文件
     * @param service 打印设备
     */
    private static void printPDF(File file, DocFlavor flavor, PrintService service) {
        try {
            PDDocument doc = PDDocument.load(file);
            PDFRenderer renderer = new PDFRenderer(doc);
            int pageCount = doc.getNumberOfPages();
            for(int i=0;i<pageCount;i++){
                File f = new File(file.getParent() + File.separator + "temp_" + i + ".png");
                BufferedImage image = renderer.renderImageWithDPI(i, 96);
                ImageIO.write(image, "PNG", f);
                normalPrint(f, flavor, LANDSCAPE, service);
                f.delete();
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    /**
     * 打印机打印Word
     * @param filepath 打印文件路径
     * @param deviceName 传入哪个设备名称，用哪个设备打印
     */
    private static void printWord(String filepath, String deviceName) {
        if(filepath.isEmpty()){
            return;
        }
        ComThread.InitSTA();
        //使用Jacob创建 ActiveX部件对象：
        ActiveXComponent word=new ActiveXComponent("Word.Application");
        //打开Word文档
        Dispatch doc=null;
        Dispatch.put(word, "Visible", new Variant(false));
        word.setProperty("ActivePrinter", new Variant(deviceName));
        Dispatch docs=word.getProperty("Documents").toDispatch();
        doc=Dispatch.call(docs, "Open", filepath).toDispatch();
        try {
            Dispatch.call(doc, "PrintOut");//打印
        } catch (Exception e) {
            e.printStackTrace();
        }finally{
            try {
                if(doc!=null){
                    //关闭文档
                    Dispatch.call(doc, "Close",new Variant(0));
                }
            } catch (Exception e2) {
                e2.printStackTrace();
            }
            word.invoke("Quit", new Variant[] {});//关闭进程
            //释放资源
            ComThread.Release();
        }
    }

    /**
     * 打印Excel
     * @param filePath 打印文件路径，形如 E:\\temp\\tempfile\\1494607000581.xls
     * @param deviceName 传入哪个设备名称，用哪个设备打印
     */
    private static void printExcel(String filePath, String deviceName){
        if(filePath.isEmpty()){
            return;
        }
        ComThread.InitSTA();

        ActiveXComponent xl=new ActiveXComponent("Excel.Application");
        try {
            Dispatch.put(xl, "Visible", new Variant(true));
            Dispatch workbooks = xl.getProperty("Workbooks").toDispatch();
            Dispatch excel=Dispatch.call(workbooks, "Open", filePath).toDispatch();
            Dispatch.callN(excel,"PrintOut",new Object[]{Variant.VT_MISSING, Variant.VT_MISSING, new Integer(1),
                    new Boolean(false), deviceName, new Boolean(true),Variant.VT_MISSING, ""});
            Dispatch.call(excel, "Close", new Variant(false));
        } catch (Exception e) {
            e.printStackTrace();
        } finally{
            xl.invoke("Quit",new Variant[0]);
            ComThread.Release();
        }
    }

    /**
     * 打印PPT
     * @param filePath
     * @param deviceName
     */
    private static void printPPT(String filePath, String deviceName) {
        File file = new File(filePath);
        File pdfFile = new File(file.getParentFile().getAbsolutePath() + file.getName() + ".pdf");
        ActiveXComponent app = null;
        Dispatch ppt = null;
        try {
            ComThread.InitSTA();
            app = new ActiveXComponent("PowerPoint.Application");
            Dispatch ppts = app.getProperty("Presentations").toDispatch();

            ppt = Dispatch.call(ppts, "Open", filePath, true, true, false).toDispatch();
            Dispatch.call(ppt, "SaveAs", pdfFile.getAbsolutePath(), 32);

        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            if (ppt != null) {
                Dispatch.call(ppt, "Close");
            }
            if (app != null) {
                app.invoke("Quit");
            }
            ComThread.Release();
        }
        try {
            print(pdfFile.getAbsolutePath(), deviceName);
            pdfFile.delete();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * 接口，在打印结束后调用
     */
    public interface AfterPrint {
        void run();
    }
}
