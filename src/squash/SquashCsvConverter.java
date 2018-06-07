package squash;

import javax.swing.SwingUtilities;
import javax.swing.UIManager;

import components.FileChooser;

public class SquashCsvConverter {

	public static void main(String[] args) {
		SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                //Turn off metal's use of bold fonts
                UIManager.put("swing.boldMetal", Boolean.FALSE); 
                FileChooser fc = new FileChooser();
                fc.createAndShowGUI();
            }
        });
		/*File csv = new File(args[0]);
		//System.out.println(csv.getPath());
		CsvUtils csvUtils = new CsvUtils(csv);
		try {
			csvUtils.csvToFormattedXlsx();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}*/
	}
}
