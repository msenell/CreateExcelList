package mainPack;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.Image;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.image.BufferedImage;
import java.awt.image.ImageObserver;
import java.awt.image.ImageProducer;
import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;

import javax.imageio.ImageIO;
import javax.swing.Box;
import javax.swing.GroupLayout;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.UIManager;

import jxl.Sheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MainFrame extends JFrame 
{
	JPanel pnlMain; //Ana panel ve ProgressBar Paneli
	JTextField txtDirPath, txtFileName; //Excel dosyalarýný içeren klasör yolu ve çýktý dosya adý.
	public JCheckBox chkSize, chkBorder; //Hücre otomatik boyutlandýrma ve kenarlýk seçeneði.
	JButton btnStart; // Ýþlem baþlatma butonu
	JLabel lblDirPath, lblFileName;
	Box bxDirPath, bxFileName, bxAllVert = null;
	GroupLayout lytGroup;
	ExcelFunctions ef;
	
	private void setLookAndFeel() //Dosya seçiçi görünümünü Windows'a uyarlar.
	{
		try {
		UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
		}
		catch ( Exception e ) {
		System.err.println( "Could not use Look and Feel:" + e );
		}
	}
	
	public String SelectATxtDir()  //TxT Dosyasý Seçmemizi Saðlayan Method.
	{
		setLookAndFeel(); 
		JFileChooser tc = new JFileChooser("C:\\");
		tc.setDialogTitle("Klasör Seçiniz");
	    tc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		int result = tc.showOpenDialog(null);
		if(result == JFileChooser.APPROVE_OPTION) //Secim onaylandý ise
		{
			return tc.getSelectedFile().toString(); //Seçilen Dosyanýn Yolunu Döndürür.
		}
		else if(result == JFileChooser.CANCEL_OPTION) //Ýptal edildi ise
		{
			System.out.println("Islem iptal edildi!");
		}
		else 
		{
			System.out.println("Bir hata oluþtu!");
		}
		
		return null;
	}

	private void createOutDir(File outDir)
	{
		if(outDir.exists()) //Klasör varmý diye kontrol ettiriyoruz.
		{
			
		}
		else
		{
			try 
			{
				if(outDir.mkdir()) //Buradaki mkdir klasörü belirtir.
					{
			
					}
				else
				{
					
				}
			} catch (Exception e) 
			{
				
			}
		}
		
	}
	
	private void writeAll()
	{
		
		File dir = new File(txtDirPath.getText());
		File outDir = new File(dir.getAbsolutePath() + "\\Output");
		createOutDir(outDir);
		File lst = new File(outDir.getAbsoluteFile() + "\\" + txtFileName.getText() + ".xls");
		
		FilenameFilter fileNameFilter = new FilenameFilter() {
			   
            @Override
            public boolean accept(File dir, String name) {
               if(name.lastIndexOf('.')>0)
               {
                  // get last index for '.' char
                  int lastIndex = name.lastIndexOf('.');
                  
                  // get extension
                  String str = name.substring(lastIndex);
                  
                  // match path name extension
                  if(str.equals(".xls"))
                  {
                     return true;
                  }
               }
               return false;
            }
         };
         File[] list=dir.listFiles(fileNameFilter);
         
		WritableWorkbook lstWB;
		Sheet s1;
		String[] row = new String[7];
		
		row[0] = "No";
		row[1] = "Müþteri";
		row[2] = "Þasi No";
		row[3] = "Motor No";
		row[4] = "Parça";
		row[5] = "Tarih";
		row[6] = "Kleym No";
		
		lstWB = ef.createWrite(lst);
		ef.write2Excel(lstWB, row, 1);

		for(int i=0; i<list.length; i++)
		{
			s1 = ef.getSheet(list[i],dir.getAbsolutePath());
			row = ef.getARow(s1);
			ef.write2Excel(lstWB, row, i+2);
		}

		if(ef.chkSize)
			ef.cellFormatting(lstWB);
		
		try {
			lstWB.write();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			lstWB.close();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	
	public MainFrame() //Kurucu metod
	{
		JLabel lblSign = new JLabel("Mustafa ÞENEL 2014");
		JPanel pnlSign = new JPanel();
		pnlMain = new JPanel();
		txtDirPath = new JTextField("Klasör Yolunu Giriniz");
		txtFileName = new JTextField("Dosya Adýný Giriniz");
		chkBorder = new JCheckBox();
		chkSize = new JCheckBox();
		btnStart = new JButton("Baþlat");
		lblDirPath = new JLabel("Dosyalarý Ýçeren Klasör Yolu : ");
		lblFileName = new JLabel("Çýktý Dosya Adý : ");
		lytGroup = new GroupLayout(pnlMain);
		ef = new ExcelFunctions();
		
		pnlMain.add(lblDirPath);
		pnlMain.add(lblFileName);
		pnlMain.add(txtDirPath);
		pnlMain.add(txtFileName);
		pnlMain.add(chkSize);
		pnlMain.add(chkBorder);
		pnlMain.add(btnStart);
		
		
		pnlMain.setLayout(lytGroup);
		
		lytGroup.setAutoCreateGaps(true);
		lytGroup.setAutoCreateContainerGaps(true);
		
		lytGroup.setHorizontalGroup(
				   lytGroup.createSequentialGroup()
				  
				   .addGroup(lytGroup.createSequentialGroup()
					         .addGroup(lytGroup.createParallelGroup(GroupLayout.Alignment.LEADING)
					        		 .addComponent(lblDirPath)
					        		 .addComponent(lblFileName))
					         .addGroup(lytGroup.createParallelGroup(GroupLayout.Alignment.LEADING)
					        		 .addComponent(txtDirPath)
					        		 .addComponent(txtFileName)
					        		 .addComponent(btnStart))
					         .addGroup(lytGroup.createParallelGroup(GroupLayout.Alignment.LEADING)
					        		 .addComponent(chkSize)
					        		 .addComponent(chkBorder))
					        		 
					));
				
				      
		lytGroup.setVerticalGroup(lytGroup.createSequentialGroup()
			    .addGroup(lytGroup.createParallelGroup(GroupLayout.Alignment.BASELINE)
			        .addComponent(lblDirPath)
			        .addComponent(txtDirPath)
			        .addComponent(chkSize))
			    
			    .addGroup(lytGroup.createSequentialGroup()
			            .addGroup(lytGroup.createParallelGroup(GroupLayout.Alignment.BASELINE)
			                .addComponent(lblFileName)
			                .addComponent(txtFileName)
			                .addComponent(chkBorder))
			    .addGroup(lytGroup.createParallelGroup(GroupLayout.Alignment.BASELINE)
			    		.addComponent(btnStart))
			    		));
		lblSign.setFont(new Font(Font.SANS_SERIF, Font.BOLD, 12));
		lblSign.setForeground(Color.BLUE);
		pnlSign.setBackground(Color.LIGHT_GRAY);
		pnlSign.add(lblSign, BorderLayout.CENTER);
		
		chkBorder.setText("Border");
		chkSize.setText("Cell Auto-Size");
		chkSize.setSelected(true);
		chkBorder.setSelected(true);
		
		txtDirPath.addMouseListener(new MouseAdapter(){
            @Override
            public void mouseClicked(MouseEvent e){
                txtDirPath.setText(SelectATxtDir());
            }
        });
		
		btnStart.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) 
			{
				txtDirPath.setEnabled(false);
				txtFileName.setEnabled(false);
				chkBorder.setEnabled(false);
				chkSize.setEnabled(false);
				btnStart.setEnabled(false);
				writeAll();
				txtDirPath.setEnabled(true);
				txtFileName.setEnabled(true);
				chkBorder.setEnabled(true);
				chkSize.setEnabled(true);
				btnStart.setEnabled(true);
				JOptionPane.showMessageDialog(pnlMain, "Ýþlem Tamamlandý!");
			}
		});
		
		txtFileName.addMouseListener(new MouseAdapter() 
		{
			public void mouseClicked(MouseEvent me)
			{
				txtFileName.setText("");
			}
		});
		
		chkBorder.addItemListener(new ItemListener() {

            @Override
            public void itemStateChanged(ItemEvent e) {
                if(e.getStateChange() == ItemEvent.SELECTED)
                	ef.chkBorder = true;
                else
                	ef.chkBorder = false;
                    
            }
        });
		
		chkSize.addItemListener(new ItemListener() {

            @Override
            public void itemStateChanged(ItemEvent e) {
                if(e.getStateChange() == ItemEvent.SELECTED)
                	ef.chkSize = true;
                else
                	ef.chkSize = false;
                    
            }
        });
		pnlMain.setBackground(Color.WHITE);
		
		this.setLayout(new BorderLayout());
		this.add(pnlMain, BorderLayout.CENTER);
		this.add(pnlSign, BorderLayout.SOUTH);
		this.setTitle("Excel List Creator");
		this.setSize(500, 150);
		this.setResizable(false);
		this.setLocationRelativeTo(null);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		this.setVisible(true);
		
		
	}
	
	public static void main(String[] args) 
	{
		new MainFrame();

	}

}
