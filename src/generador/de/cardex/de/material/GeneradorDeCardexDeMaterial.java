/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package generador.de.cardex.de.material;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 *
 * @author Manuel
 */
public class GeneradorDeCardexDeMaterial {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream("PLANTILLA CARDEX MATERIAL.xls"));
        HSSFWorkbook plantilla = new HSSFWorkbook(fs);
        HSSFSheet hoja1 = plantilla.getSheetAt(0);
        int posicion = 0;
        
        String rutaSOGECOMA="";
        JFileChooser fileChooser = new JFileChooser();              
        int result = fileChooser.showOpenDialog(null);  
        if ( result == JFileChooser.APPROVE_OPTION ){            
            rutaSOGECOMA = fileChooser.getSelectedFile().getAbsolutePath(); 
        }
        POIFSFileSystem fs1 = new POIFSFileSystem(new FileInputStream(rutaSOGECOMA));
        HSSFWorkbook sogecoma = new HSSFWorkbook(fs1);
        HSSFSheet materiales = sogecoma.getSheetAt(3);
        int numMats = materiales.getLastRowNum();
        HSSFCellStyle estilo = plantilla.createCellStyle();
        HSSFFont fuente=plantilla.createFont();
        fuente.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        estilo.setFont(fuente);
        
        
        for (int m=1;m<=numMats;m++){
            try {
                hoja1.getRow(2).createCell(1).setCellValue(materiales.getRow(m).getCell(1).getStringCellValue());
                hoja1.getRow(2).getCell(1).setCellStyle(estilo);
                hoja1.getRow(3).createCell(1).setCellValue(materiales.getRow(m).getCell(2).getStringCellValue());
                hoja1.getRow(3).getCell(1).setCellStyle(estilo);
                FileOutputStream elFichero = new FileOutputStream("CARDEX "+materiales.getRow(m).getCell(1).getStringCellValue()+".xls");
                plantilla.write(elFichero);
                elFichero.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}
