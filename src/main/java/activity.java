import java.awt.AWTException;
import java.awt.Color;
import java.awt.Container;
import java.awt.Frame;
import java.awt.Image;
import java.awt.SystemTray;
import java.awt.TextArea;
import java.awt.Toolkit;
import java.awt.TrayIcon;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.security.GeneralSecurityException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Timer;
import java.util.TimerTask;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.NoSuchProviderException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

import com.sun.mail.util.MailSSLSocketFactory;
public class activity {
    public static String xlsUrl="测试同步！";
    public static String pictureUrl="D:\\Desktop\\发送邮件\\新建位图图像.bmp";
    public static String sendEamil= "guofan@accbio.com.cn,664130988@qq.com";
    public static String myEamil= "guofan@accbio.com.cn";
    public static String EamilTitle= "";
    public static String EamilContent= "";
    public static PrintStream printStream;
    public static Statement statement=null;
    public static Connection ct=null;
    public static Session session=null;
    public static Transport transport =null;
    public static Timer timer;
    public static JTextField TSend;
    public static JTextArea  TvContent,TvSqlTab;
    public static JTextField  TvAccessory;
    public static JTextField TTitle;
    public static TextArea  TvLog;
    public static String currentDate,currentDate1;
    public static List<String[]> src;
    public static SimpleDateFormat format,format1;
    public static void main(String[] args) {
        initActivity();

    }
    //初始化界面
    public static void  initActivity(){
        final JFrame f=new JFrame("自动发送邮件系统");//主界面
        //创建第一个table主页
        JTabbedPane tab1=new JTabbedPane();//tab1
        tab1.setBounds(0, 0, 410, 660);
        JPanel group=new JPanel();
        group.setLayout(null);
        tab1.add(group,"主页");
        f.add(tab1);
        //创建第二个table设置
        JPanel group2=new JPanel();
        group2.setLayout(null);
        JLabel setTable2=new JLabel();
        setTable2.setText("设置执行计划：");
        setTable2.setBounds(10, 0, 150, 20);
        JLabel labLine2=new JLabel();
        labLine2.setBounds(5,20,400,100);
        labLine2.setBorder(BorderFactory.createLineBorder(Color.DARK_GRAY));
        JLabel setTable3=new JLabel();
        setTable3.setText("设置发送附件数据：");
        setTable3.setBounds(5, 120, 150, 20);
        JLabel labLine3=new JLabel();
        labLine3.setBounds(5,140,400,100);
        labLine3.setBorder(BorderFactory.createLineBorder(Color.DARK_GRAY));
        group2.add(setTable2);
        group2.add(labLine2);
        group2.add(labLine3);
        group2.add(setTable3);
        tab1.add(group2,"设置");
        f.add(tab1);
        //创建第三个table关于
        JPanel group3=new JPanel();
        group3.setLayout(null);
        JLabel versions=new JLabel();
        versions.setText("Versions: 1.13 ");
        versions.setBounds(10, 0, 400, 70);
        JLabel author=new JLabel();
        author.setText("Author: GuoFan");
        author.setBounds(10,30,400,70);
        JLabel contactInformation=new JLabel();
        contactInformation.setText("Contact information: 664130988@qq.com");
        contactInformation.setBounds(10,60,400,70);
        JLabel labLine=new JLabel();
        labLine.setBounds(5,20,400,100);
        labLine.setBorder(BorderFactory.createLineBorder(Color.DARK_GRAY));
        group3.add(versions);
        group3.add(author);
        group3.add(contactInformation);
        group3.add(labLine);
        JLabel title=new JLabel();
        title.setText("About：");
        title.setBounds(10,0,100,20);
        group3.add(title);
        tab1.add(group3,"关于");
        f.add(tab1);
        TSend=new JTextField(25);
        TTitle=new JTextField(25);
        TvContent=new JTextArea ();
        TvAccessory=new JTextField ();
        TvSqlTab=new JTextArea();
        final JTextField  TvRecyDate=new JTextField ();
        final JTextField  TvRecyDate_Time=new JTextField ();
        TvLog=new TextArea ();
        f.setLayout(null);
        f.setVisible(true);
        f.setSize(420, 690);
        f.setLocation(450,100);
        f.setResizable(false);
        JLabel label=new JLabel("收件人：");
        JLabel labe2=new JLabel("标题：");
        group.add(label);
        label.setBounds(10, 10, 60, 30);
        group.add(TSend);
        TSend.setBounds(70, 10, 330, 30);
        group.add(labe2);
        labe2.setBounds(10, 50, 60, 30);
        group.add(TTitle);
        TTitle.setBounds(70, 50, 330, 30);
        JLabel labLine0=new JLabel();
        labLine0.setBounds(5,0,400,90);
        labLine0.setBorder(BorderFactory.createLineBorder(Color.DARK_GRAY));
        group.add(labLine0);
        JLabel labeContent=new JLabel("主题：");
        group.add(labeContent);
        labeContent.setBounds(10, 90, 45, 30);
        TvContent.setLineWrap(true);
        group.add(TvContent);
        TvContent.setBounds(10, 120,390, 190);
        JLabel labLine01=new JLabel();
        labLine01.setBounds(5,115,400,200);
        labLine01.setBorder(BorderFactory.createLineBorder(Color.DARK_GRAY));
        group.add(labLine01);
        //要发送的表名
        JLabel SqlTable=new JLabel("<html><body><p align=\"center\">数据库表名<br/>Excel sheet名<br/>格式：表名,sheet名;</p></body></html>");
        group2.add(SqlTable);
        SqlTable.setBounds(10, 150, 80, 90);
        group2.add(TvSqlTab);
        TvSqlTab.setLineWrap(true);
        TvSqlTab.setBounds(100, 145, 300, 90);
        JLabel labeAccessory=new JLabel("附件地址：");
        group.add(labeAccessory);
        labeAccessory.setBounds(10, 325, 70, 30);
        group.add(TvAccessory);
        TvAccessory.setBounds(80, 325, 320, 30);
        JLabel labLine02=new JLabel();
        labLine02.setBounds(5,320,400,80);
        labLine02.setBorder(BorderFactory.createLineBorder(Color.DARK_GRAY));
        group.add(labLine02);
        //设置TAb2执行计划
        JLabel labeRecyDate=new JLabel("设置为每");
        group2.add(labeRecyDate);
        labeRecyDate.setBounds(10, 25, 100, 30);
        TvRecyDate.setText("01");
        TvRecyDate.setBounds(115, 25, 30, 30);
        JLabel NumDate=new JLabel("第");
        NumDate.setBounds(100, 25, 100, 30);
        group2.add(NumDate);
        group2.add(TvRecyDate);
        //设置循环循环周期
        final JComboBox cmb=new JComboBox();
        cmb.addItem("天");
        cmb.addItem("周");
        cmb.addItem("月");
        cmb.addItem("年");
        cmb.setBounds(60, 25, 40, 30);
        group2.add(cmb);
        JLabel labeRecyDate_day=new JLabel("天");
        group2.add(labeRecyDate_day);
        labeRecyDate_day.setBounds(145, 25, 20, 30);
        TvRecyDate_Time.setText("23:00:00");
        group2.add(TvRecyDate_Time);
        TvRecyDate_Time.setBounds(160, 25, 60, 30);
        JLabel labeRecyDate_Time=new JLabel("时间发送!");
        group2.add(labeRecyDate_Time);
        labeRecyDate_Time.setBounds(230, 25, 60, 30);
        final JButton btSetDate=new JButton("插入执行计划");
        group2.add(btSetDate);
        btSetDate.setBounds(200, 75, 120, 30);
        //设置清除所有执行继续
        final JButton btClearDate=new JButton("删除执行计划");
        group2.add(btClearDate);
        btClearDate.setBounds(50, 75, 120, 30);
        btClearDate.addActionListener(new ActionListener() {
            //删除执行计划按钮
            @Override
            public void actionPerformed(ActionEvent e) {
                if(clearSQL())
                {
                    JOptionPane.showMessageDialog(f,"执行计划已删除！","提示信息",JOptionPane.WARNING_MESSAGE);
                }else{
                    JOptionPane.showMessageDialog(f,"执行计划已删除失败！","提示信息",JOptionPane.WARNING_MESSAGE);
                }
            }
        });
        //设置发送数据
        //按钮 测试、启动、取消
        final JButton btTest=new JButton("发送测试");
        group.add(btTest);
        btTest.setBounds(10, 360, 100, 30);
        final JButton btStart=new JButton("启动");
        group.add(btStart);
        btStart.setBounds(150, 360, 100, 30);
        JButton btEnd=new JButton("取消");
        group.add(btEnd);
        btEnd.setBounds(300, 360, 100, 30);
        //调试信息
        JLabel labeLog=new JLabel("调试信息：");
        group.add(labeLog);
        labeLog.setBounds(10, 400, 70, 30);
        group.add(TvLog);
        TvLog.setBounds(10, 430,390, 190);
        TvLog.setEditable(false);
        //创建附加列表
        format=new SimpleDateFormat("yyyy-MM-dd");
        format1=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        f.addWindowListener(new WindowAdapter(){
            public void windowClosing(WindowEvent e)
            {
                System.out.println("关闭程序！");
                try
                {
                    if(statement!=null && ct!=null && transport!=null){
                        statement.close();
                        ct.close();
                        transport.close();
                        System.out.println("关闭数据库连接");
                    }
                } catch (SQLException | MessagingException e1) {
                    e1.printStackTrace();

                }
                System.exit(0);
            }
            public void windowActivated(WindowEvent e)
            {
                System.out.println("激活窗口！");

            }
            //托盘化窗口
            public void windowIconified(WindowEvent e) {
                f.dispose(); //窗口最小化时dispose该窗口
            }
        });
        //设置软件在托盘上显示的图标
        Toolkit tk = Toolkit.getDefaultToolkit();
        Image img = tk.getImage("emailico.jpg");//*.gif与该类文件同一目录
        SystemTray systemTray = SystemTray.getSystemTray(); //获得系统托盘的实例
        TrayIcon trayIcon = null;
        try {
            trayIcon = new TrayIcon(img, "邮件自动发送小程序");
            systemTray.add(trayIcon); //设置托盘的图标，*.gif与该类文件同一目录
            f.setIconImage(img);
            trayIcon.setImageAutoSize(true);
        } catch (AWTException e2) {
            e2.printStackTrace();
        }
        //双击托盘图标，软件正常显示
        trayIcon.addMouseListener(new MouseAdapter() {
            public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 1) //双击托盘窗口再现
                    //置此 frame 的状态。该状态表示为逐位掩码。
                    f.setExtendedState(Frame.NORMAL); //正常化状态
                f.setVisible(true);
            }
        });
        btTest.addActionListener(new ActionListener() {
            //测试发送按钮
            @Override
            public void actionPerformed(ActionEvent e) {
                sendEamil=TSend.getText().toString();
                EamilTitle=TTitle.getText().toString();
                EamilContent=TvContent.getText().toString();
                if(EamilTitle.isEmpty()||sendEamil.isEmpty())
                {
                    TvLog.append("收件人或者标题没填写！\r\n");
                    JOptionPane.showMessageDialog(f, "收件人或者标题未填写!", "提示消息", JOptionPane.WARNING_MESSAGE);

                }else{
                    currentDate=format.format(new Date(System.currentTimeMillis()));
                    if(!TvSqlTab.getText().toString().equals(""))
                    {
                        createExcel(TvSqlTab.getText().toString(),"销售周报"+currentDate+".xls");
                        xlsUrl=TvAccessory.getText().toString();
                    }
                    sendEamil(sendEamil,xlsUrl,EamilTitle,EamilContent);
                }

            }
        });
        //初始化日志文件
        File file=new File("ExceptionLog.log");
        try {
            printStream =new PrintStream(new FileOutputStream(file,true),true);
        } catch (FileNotFoundException e1) {
            e1.printStackTrace(printStream);
        }
        //启动后执行顺序
        btStart.addActionListener(new ActionListener() {

            @Override
            public void actionPerformed(ActionEvent e) {
                recycleTask();
                btStart.setEnabled(false);
                btTest.setEnabled(false);
                TTitle.setEditable(false);
                TvRecyDate.setEditable(false);
                TSend.setEditable(false);
                TvAccessory.setEditable(false);
                TvRecyDate_Time.setEditable(false);
                TvContent.setEditable(false);
                btSetDate.setEnabled(false);
                TvSqlTab.setEnabled(false);
                btClearDate.setEnabled(false);
                xlsUrl=TvAccessory.getText().toString();
                sendEamil=TSend.getText().toString();
            }
        });
        btEnd.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                btStart.setEnabled(true);
                btTest.setEnabled(true);
                TTitle.setEditable(true);
                TvRecyDate.setEditable(true);
                TSend.setEditable(true);
                TvAccessory.setEditable(true);
                TvRecyDate_Time.setEditable(true);
                TvContent.setEditable(true);
                btSetDate.setEnabled(true);
                btClearDate.setEnabled(true);
                if(timer!=null){
                    timer.cancel();
                }
            }
        });
        //插入执行计划
        btSetDate.addActionListener(new ActionListener() {
            //Boolean flag =true;
            Calendar calendar;
            Date Date ,date;
            String planDate;
            String currData;
            String str;
            String month;
            public void actionPerformed(ActionEvent e) {
                calendar=Calendar.getInstance();
                calendar.add(Calendar.MONTH, 1);
                System.out.println(calendar.toString()+"=================");
                calendar.get(Calendar.MARCH);
                str="0"+calendar.get(Calendar.MARCH);
                month = str.substring(str.length()-2,str.length());
                currData=calendar.get(Calendar.YEAR)+"-"+month+"-"+TvRecyDate.getText().toString()+" "+TvRecyDate_Time.getText().toString();
                try {
                    Date=format1.parse(currData);
                    calendar.setTime(Date);
                    System.out.println("======"+currData+format1.format(Date));
                } catch (ParseException e1) {
                    e1.printStackTrace();
                }
//				TvRecyDate.setEditable(false);
//				TvRecyDate_Time.setEditable(false);
                for(int a=0; a<52;a++)
                {
                    date=calendar.getTime();
                    planDate=format1.format(date);
                    System.out.println(a+planDate);
                    write(planDate,"未发送");
                    TvLog.append(planDate+"执行计划生成！\r\n");
                    if(cmb.getSelectedIndex()==0)
                    {
                        System.out.println("选择了第0个");
                        calendar.add(Calendar.DAY_OF_MONTH,1);
                    } else if(cmb.getSelectedIndex()==1)
                    {
                        System.out.println("选择了第1个");
                        calendar.add(Calendar.WEDNESDAY, 1);
                    }else if(cmb.getSelectedIndex()==2)
                    {
                        System.out.println("选择了第2个");
                        calendar.add(Calendar.MONTH, 1);
                    }else if(cmb.getSelectedIndex()==3)
                    {
                        System.out.println("选择了第3个");
                        calendar.add(Calendar.YEAR, 1);
                    }
                }
//				TvRecyDate.setEditable(true);
//				TvRecyDate_Time.setEditable(true);
                JOptionPane.showMessageDialog(f, "插入执行计划成功", "提示消息", JOptionPane.WARNING_MESSAGE);
            }
        });
    }
    //初始化邮件
    public static void initEamil()
    {
        Properties properties = new Properties();
        properties.setProperty("mail.host","smtp.accbio.com.cn");
        properties.setProperty("mail.transport.protocol","smtp");
        properties.setProperty("mail.smtp.auth","true");
        MailSSLSocketFactory sf=null;
        try {
            sf = new MailSSLSocketFactory();
        } catch (GeneralSecurityException e) {
            e.printStackTrace(printStream);
        }
        sf.setTrustAllHosts(true);
        properties.put("mail.smtp.ssl.enable", "true");
        properties.put("mail.smtp.ssl.socketFactory", sf);
        //创建一个session对象
        session = Session.getDefaultInstance(properties, new Authenticator() {
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(myEamil,"fanguo2019");
            }
        });
        //开启debug模式
        session.setDebug(true);
        //获取连接对象
        try {
            transport = session.getTransport();
        } catch (NoSuchProviderException e) {
            e.printStackTrace(printStream);
        }
        //连接服务器
        try {
            transport.connect("smtp.accbio.com.cn",myEamil,"fanguo2019");
        } catch (MessagingException e) {
            // TODO 自动生成的 catch 块
            e.printStackTrace(printStream);
        }
    }

    //发送邮件
    public static void sendEamil(String eamilAdd,String acctss,String titil,String content)
    {
        MimeMessage mimeMessage =null;
        if(transport==null ||session==null||!transport.isConnected())
        {
            initEamil();
        }
        try {
            mimeMessage = complexEmail(session,eamilAdd,acctss,titil,content);
        } catch (MessagingException e1) {
            e1.printStackTrace(printStream);
        }
        //发送邮件
        try {
            transport.sendMessage(mimeMessage,mimeMessage.getAllRecipients());
            transport.close();
            session=null;
            transport=null;
        } catch (MessagingException e) {
            // TODO 自动生成的 catch 块
            e.printStackTrace(printStream);
            TvLog.append("发送失败\r\n");
        }
        TvLog.append("已发送\r\n");
    }
    //连接数据库
    public static MimeMessage complexEmail(Session session,String eamilAdd,String acctss,String titil,String content) throws MessagingException {
        //消息的固定信息
        MimeMessage mimeMessage = new MimeMessage(session);
        //发件人
        mimeMessage.setFrom(new InternetAddress(myEamil));
        new InternetAddress();
        //收件人
        mimeMessage.setRecipients(Message.RecipientType.TO,InternetAddress.parse(eamilAdd));
        //邮件标题
        mimeMessage.setSubject(titil);
        //邮件内容
        //准备图片数据
        //MimeBodyPart image = new MimeBodyPart();
        //DataHandler handler = new DataHandler(new FileDataSource(pictureUrl));
        // image.setDataHandler(handler);
        //image.setContentID("test.png"); //设置图片id
        //准备文本
        MimeBodyPart text = new MimeBodyPart();
        text.setContent(content,"text/html;charset=utf-8");
        //拼装邮件正文
        MimeMultipart mimeMultipart = new MimeMultipart();
        //mimeMultipart.addBodyPart(image);
        mimeMultipart.addBodyPart(text);
        //mimeMultipart.setSubType("related");//文本和图片内嵌成功
        //将拼装好的正文内容设置为主体
        MimeBodyPart contentText = new MimeBodyPart();
        contentText.setContent(mimeMultipart);
        //拼接附件
        MimeMultipart allFile = new MimeMultipart();
        allFile.addBodyPart(contentText);//正文
        //附件
        MimeBodyPart appendix = new MimeBodyPart();
        appendix.setDataHandler(new DataHandler(new FileDataSource("销售周报"+currentDate+".xls")));
        String fileName=new File("销售周报"+currentDate+".xls").getName().toString();
        appendix.setFileName("销售周报"+currentDate+".xls");
        allFile.addBodyPart(appendix);//附件
        allFile.setSubType("mixed"); //正文和附件都存在邮件中，所有类型设置为mixed
        //放到Message消息中
        mimeMessage.setContent(allFile);
        mimeMessage.saveChanges();//保存修改
        return mimeMessage;
    }
    //周期事件
    public static void  recycleTask()
    {
        TimerTask task=new TimerTask(){
            String[] str;
            String sendDate,sendType;
            Date Date;
            String  nextDate;
            Calendar calendar;
            @Override
            public void run() {
                //获取当前时间
                calendar=Calendar.getInstance();
                calendar.getTime();
                Date=calendar.getTime();
                currentDate=format.format(Date);
                currentDate1=format1.format(Date);
                //获取数据库时间
                str=readSQL(currentDate1);
                sendDate=str[0];
                sendType=str[1];
                System.out.println("发送时间！"+sendDate+"-----现在时间："+currentDate1);
                if(sendDate.equals(currentDate1)&&sendType.equals("未发送"))
                {
                    createExcel(TvSqlTab.getText().toString(),"销售周报"+currentDate+".xls");
                    xlsUrl=TvAccessory.getText().toString();
                    sendEamil=TSend.getText().toString();
                    EamilTitle=TTitle.getText().toString();
                    EamilContent=TvContent.getText().toString();
                    if(xlsUrl.isEmpty()||sendEamil.isEmpty())
                    {
                        TvLog.append("收件人或者附加没添加！\r\n");
                    }
                    sendEamil(sendEamil,xlsUrl,EamilTitle,EamilContent);
                    update(currentDate,"已发送");
                    TvLog.append(currentDate+"已发送\r\n");
                    System.out.println("修改发送标记|"+currentDate+"已发送");
                }
            }
        };
        timer =new Timer();
        long delay=0;
        long period=1000;
        timer.scheduleAtFixedRate(task, 1000, 1000);
    }
    //连接数据库
    public static void connectSQL()
    {
        //1、加载驱动器
        try {
            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
        } catch (ClassNotFoundException e) {
            // TODO 自动生成的 catch 块
            e.printStackTrace();
        }
        try {
            if(ct==null)
            {
                //ct=DriverManager.getConnection("jdbc:sqlserver://qds168257330.my3w.com:1433;databaseName=qds168257330_db","qds168257330","guofan6889168");
                ct=DriverManager.getConnection("jdbc:sqlserver://127.0.0.1:7081;databaseName=ecology","guofan","guofan123456");
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        //创建发送端
        System.out.println("链接成功！");
        //4、构造一个Statement对象,用来发送SQL的载体
        try {
            if(statement==null)
            {
                statement=ct.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    //读取数据库
    public static String[] readSQL(String currDate)
    {
        String[] str=new String[2];
        String query="  select top 1 date ,type from sendLog where date>='"+currDate.toString()+"' order by id asc";
        if(statement==null||ct==null)
        {
            connectSQL();
        }
        try {
            ct.prepareStatement(query,ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
            ResultSet res=statement.executeQuery(query);
            res.absolute(1) ;
            str[0]=res.getString("date").toString();
            str[1]=res.getString("type").toString();
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return str;
    }
    //清除记录
    public static Boolean clearSQL()
    {
        String query="delete from sendLog";
        if(statement==null||ct==null)
        {
            connectSQL();
        }
        try {
            int a=statement.executeUpdate(query);
            System.out.println("影响了"+a+"行！");
        } catch (SQLException e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }
    //写入数据库
    public static int write(String date,String type)
    {
        int a = 0;
        String SQL="insert into sendLog(date,type) values ( '"+date+"','"+type+"')";
        //5、发送SQL
        try
        {
            if(statement==null||ct==null)
            {
                connectSQL();
            }
            a=statement.executeUpdate(SQL);
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return a;
    }
    //修改数据库
    public static int update(String date,String type)
    {
        int a = 0;
        String SQL="update sendLog set type ='"+type+"' where date='"+date+"'";
        //5、发送SQL
        try
        {
            if(statement==null||ct==null)
            {
                connectSQL();
            }
            a=statement.executeUpdate(SQL);
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return a;
    }
    //下载一个表并生成excel
    public static void createExcel(String TableName,String ExcelPath)
    {
        //======写入Excel==================================================
        //1、创建一个excel
        File writeXls=new File(ExcelPath);
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(writeXls);
        } catch (FileNotFoundException e1) {
            // TODO 自动生成的 catch 块
            e1.printStackTrace();
        }
        //3、创建创建表对象
        HSSFWorkbook book =new HSSFWorkbook();
        //2、创建流对象
        String[] stre=TableName.split(";");
        String[] child=new String[2];
        //%号处理函数
        NumberFormat nf=NumberFormat.getPercentInstance();
        //创建会计格式
        HSSFCellStyle cellStyle = book.createCellStyle();
        HSSFDataFormat format= book.createDataFormat();
        cellStyle.setDataFormat(format.getFormat("_ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ ;_ ¥* \"-\"??_ ;_ @_ "));
        //设置为文本
        HSSFCellStyle cellStyleText = book.createCellStyle();
        HSSFDataFormat format2= book.createDataFormat();
        cellStyleText.setDataFormat(format2.getFormat("@"));
        //设置为数值
        HSSFCellStyle cellStyleInt = book.createCellStyle();
        HSSFDataFormat format3= book.createDataFormat();
        cellStyleInt.setDataFormat(format3.getFormat("0"));
        //创建会计格式
        HSSFCellStyle cellStyleB = book.createCellStyle();
        HSSFDataFormat format4= book.createDataFormat();
        cellStyleB.setDataFormat(format.getFormat("_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * \"-\"??_ ;_ @_ "));
        for(int a=0;a<stre.length;a++)
        {
            child=stre[a].split(",");
            String query="SELECT * FROM  "+child[0].toString();
            if(statement==null||ct==null)
            {
                connectSQL();
            }
            try {
                //4、创建表
                HSSFSheet mSheet=book.createSheet(child[1].toString());
                //冻结首行
                mSheet.createFreezePane( 0, 1, 0, 1 );
                ct.prepareStatement(query,ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
                ResultSet res=statement.executeQuery(query);
                int columnCount=res.getMetaData().getColumnCount();
                String str=null;
                HSSFRow readrow;
                int row=0;
                String ColumnName="";
                System.out.println("创建第"+a+"个表");
                printStream.append("创建第"+a+"个表\r\n");
                while(res.next())
                {
                    readrow =mSheet.createRow(row);//创建行，从1开始
                    System.out.println("创建第"+row+"行");
                    printStream.append("创建第"+row+"行\r\n");
                    for(int i=1;i<=columnCount;i++)
                    {
                        ColumnName=res.getMetaData().getColumnName(i);

                        if(row==0)
                        {
                            str=res.getMetaData().getColumnName(i);
                            //创建一个表格的style
                            HSSFCellStyle style = book.createCellStyle();
                            style.setFillForegroundColor(HSSFColor.GREY_50_PERCENT.index);
                            style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                            style.setFillBackgroundColor(HSSFColor.WHITE.index);
                            //生成一个字体
                            HSSFFont font = book.createFont();
                            font.setFontHeightInPoints((short) 10);
                            font.setColor(HSSFColor.WHITE.index);
                            font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
                            font.setFontName("宋体");
                            // 把字体 应用到当前样式
                            style.setFont(font);
                            HSSFCell readcel1=readrow.createCell(i-1);
                            readcel1.setCellStyle(style);
                            readcel1.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                            readcel1.setCellValue(str);//设置单元格内容
                            //设置列宽
                            mSheet.setColumnWidth(i, 300*str.length()+256*10);
                        }else{
                            //第二行开始
                            str=res.getString(i);
                            HSSFCell readcel1=readrow.createCell(i-1);
                            //str.matches("-?[0-9]+\\.?[0-9]*")
                            if(ColumnName.contains("(元)"))
                            {
                                //如果出现数字就设置为会计格式
                                readcel1.setCellStyle(cellStyle);
                                readcel1.setCellValue(Double.parseDouble(str.toString()));//设置单元格内容

                            }else if(ColumnName.contains("率"))
                            {
                                //如果出现数字就设置为会计格式
                                readcel1.setCellStyle(cellStyleB);
                                readcel1.setCellValue((Double)nf.parse(str.toString()).doubleValue());//设置单元格内容
                            }
                            else if(ColumnName.contains("(个)") )
                            {
                                readcel1.setCellStyle(cellStyleInt);
                                readcel1.setCellValue(Double.parseDouble(str.toString()));//设置单元格内容
                            }
                            else{

                                readcel1.setCellStyle(cellStyleText);
                                readcel1.setCellValue(str);//设置单元格内容

                            }

                        }
                    }
                    row++;

                }

            } catch (SQLException | ParseException   e) {
                e.printStackTrace(printStream);
            }


        }
        try {
            book.write(fos);
            fos.flush();
            TvAccessory.setText(writeXls.getPath());
            fos.close();

        } catch (IOException e) {
            // TODO 自动生成的 catch 块
            e.printStackTrace(printStream);
        }

    }

}
