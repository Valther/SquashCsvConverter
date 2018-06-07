/*
 * Copyright (c) 1995, 2008, Oracle and/or its affiliates. All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions
 * are met:
 *
 *   - Redistributions of source code must retain the above copyright
 *     notice, this list of conditions and the following disclaimer.
 *
 *   - Redistributions in binary form must reproduce the above copyright
 *     notice, this list of conditions and the following disclaimer in the
 *     documentation and/or other materials provided with the distribution.
 *
 *   - Neither the name of Oracle or the names of its
 *     contributors may be used to endorse or promote products derived
 *     from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS
 * IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
 * THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
 * PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL THE COPYRIGHT OWNER OR
 * CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
 * EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
 * PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
 * PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
 * NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */ 

package components;

import java.io.*;
import java.awt.*;
import java.awt.dnd.DropTarget;
import java.awt.event.*;
import javax.swing.*;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.xml.bind.JAXBException;

import squash.utils.CsvUtils;
import squash.utils.SquashReportUtils;

import java.util.Scanner;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

import utils.DragDropListener;

/*
 * FileChooserDemo.java uses these files:
 *   images/Open16.gif
 */
public class FileChooser extends JPanel
                             implements ActionListener {
    /**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private final String newline = "\n";
    JButton openButton, launchButton;
    JTextArea log;
    JFileChooser fc;
    JProgressBar progressBar;
    StringBuilder fileContent;
    File file;

	public FileChooser() {
        super(new BorderLayout());
        
        //Create the log first, because the action listeners
        //need to refer to it.
        log = new JTextArea(5,20);
        log.setMargin(new Insets(5,5,5,5));
        log.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(log);

        //Create a file chooser
        fc = new JFileChooser();
        FileFilter htmlFilter = new FileNameExtensionFilter("Fichier CSV", "csv");
        fc.setFileFilter(htmlFilter);
        fc.setAcceptAllFileFilterUsed(false);
        
        /*FileFilter docxFilter = new FileNameExtensionFilter("Document word étendu(*.docx)", "docx", "docx");
        fc.setFileFilter(docxFilter);*/
        
        openButton = new JButton("Ouvrir...",
                                 createImageIcon("images/Open16.gif"));
        openButton.addActionListener(this);
        

        //Create the save button.  We use the image from the JLF
        //Graphics Repository (but we extracted it from the jar).
        launchButton = new JButton("Lancer le traitement");
        launchButton.addActionListener(this);
        
        //For layout purposes, put the buttons in a separate panel
        JPanel buttonPanel = new JPanel(); //use FlowLayout
        String instructions = "<html><div style='text-align: center;'>"
        		+"Ouvrir/glisser l'export CSV ici</div></html>";
        JLabel instructionsLabel = new JLabel(instructions);
        progressBar = new JProgressBar();
        instructionsLabel.setHorizontalAlignment(JLabel.CENTER);
        buttonPanel.add(openButton);
        buttonPanel.add(launchButton);
        
        JPanel footerPanel = new JPanel();
        footerPanel.setLayout(new BoxLayout(footerPanel, BoxLayout.Y_AXIS));
        footerPanel.add(logScrollPane);
        footerPanel.add(progressBar, BorderLayout.PAGE_END);
        //Add the buttons and the log to this panel.
        add(buttonPanel, BorderLayout.PAGE_START);
        add(instructionsLabel, BorderLayout.CENTER);
        add(footerPanel, BorderLayout.PAGE_END);
        DragDropListener dragDropListener = new DragDropListener(this);
        new DropTarget(openButton, dragDropListener);
        new DropTarget(instructionsLabel, dragDropListener);
    }

    public void actionPerformed(ActionEvent e) {

        //Handle open button action.
        if (e.getSource() == openButton) {
            int returnVal = fc.showOpenDialog(FileChooser.this);

            if (returnVal == JFileChooser.APPROVE_OPTION) {
                handleFile(fc.getSelectedFile());
            }
            log.setCaretPosition(log.getDocument().getLength());

        //Handle save button action.
        } else if (e.getSource() == launchButton) {
        	try {
				Callable<Boolean> task = () -> {
					return fileTreatmnent();
				};
				ExecutorService executor = Executors.newFixedThreadPool(2);
				@SuppressWarnings("unused")
				Future<Boolean> future = executor.submit(task);
        		
				} catch (Exception e1) {
					e1.printStackTrace();
					JOptionPane.showMessageDialog(this, "Erreur lors du traitement du fichier:\n" + e1.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
					progressBar.setIndeterminate(false);
				}
        	
            log.setCaretPosition(log.getDocument().getLength());
        }
    }
    
    private void setCanInteract(boolean canInteract) {
    	launchButton.setEnabled(canInteract);
    	openButton.setEnabled(canInteract);
    }
    
    public void handleFile(File f) {
    	try {
    		file = f;
    		String fileName = file.getName();
            if(fileName.endsWith("csv")){
            	Scanner scanner = new Scanner(f);
    			fileContent = new StringBuilder();
                try {
                	//log.append("Opening: " + file.getName() + "." + newline);
    				while (scanner.hasNext()) {
    					fileContent.append(scanner.next());
    				}
    			} catch (Exception e1) {
    				e1.printStackTrace();
    				JOptionPane.showMessageDialog(this, "Erreur lors de la lecture du fichier: " + e1.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
    			} finally {
    				scanner.close();
    			}
                log.setText(null);
                log.append("Document squash à traiter: " + file.getPath() + newline);
            } else {
                displayInvalidFileFormat();
            }	
        } catch (Exception ex) {
        	ex.printStackTrace();
        	JOptionPane.showMessageDialog(this, "Erreur lors de la lecture du fichier: " + ex.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
        	//log.append("Erreur lors de la lecture du fichier: " + ex.getMessage() + newline);
        }
    }
    
    /**
     * Traitement du fichier passé à l'interface
     * @return
     * @throws IOException
     * @throws Docx4JException
     * @throws JAXBException
     */
    public boolean fileTreatmnent() throws IOException, JAXBException {
    	try{
    		if(file != null && fileContent != null){
    			setCanInteract(false);
        		log.append("Traitement en cours..." + newline);
        		progressBar.setIndeterminate(true);
        		CsvUtils csvUtils = new CsvUtils(file);
        		File resultFile = csvUtils.csvToFormattedXlsx();
        		log.append("Enregistrement du fichier Excel en cours à l'adresse: " + resultFile.getPath() + newline);
        		progressBar.setIndeterminate(false);
        		JOptionPane.showMessageDialog(this, "Fichier traité avec succès" + newline);
        		log.append("Images générées: "+SquashReportUtils.getNbImageGenerated() + newline);
        		log.append("Nombre de liens OK: "+SquashReportUtils.getNbHyperlinkDone() + newline);
        		log.append("Nombre de liens KO: "+SquashReportUtils.getNbHyperlinkKO() + newline);
        		log.append("Fin du traitement" + newline);
        		setCanInteract(true);
        	}
        	return true;
    	} catch(Exception ex){
    		ex.getStackTrace();
    		progressBar.setIndeterminate(false);
    		JOptionPane.showMessageDialog(this, "Erreur lors du traitement du fichier:\n" + ex.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
    	}
    	setCanInteract(true);
    	return false;
    }
    
    public void displayInvalidFileFormat() {
    	JOptionPane.showMessageDialog(this, "Le fichier doit être au format csv", "Erreur", JOptionPane.ERROR_MESSAGE);
    }

    /** Returns an ImageIcon, or null if the path was invalid. */
    protected static ImageIcon createImageIcon(String path) {
        java.net.URL imgURL = FileChooser.class.getResource(path);
        if (imgURL != null) {
            return new ImageIcon(imgURL);
        } else {
            System.err.println("Couldn't find file: " + path);
            return null;
        }
    }

    /**
     * Create the GUI and show it.  For thread safety,
     * this method should be invoked from the
     * event dispatch thread.
     */
    public void createAndShowGUI() {
        //Create and set up the window.
    	ImageIcon icon = createImageIcon("images/excel.png");
        JFrame frame = new JFrame("Squash CSV Converter");
        frame.setIconImage(icon.getImage());
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        Dimension dim = new Dimension(600, 400);
        frame.setPreferredSize(dim);
        frame.setMinimumSize(dim);
        //Add content to the window.
        frame.add(new FileChooser());

        //Display the window.
        frame.pack();
        frame.setVisible(true);
    }
}
