package TellerTotaalOntkleeft;

public class Main{
    public static void main(String[] args) {
        if (args.length == 0) {
            System.err.println("File name not specified.");
            System.exit(1);
        }

        try {
            XLSX XLSXFile = new XLSX(args[0]);
            XLSXFile.selectSheet();
            XLSXFile.countPerDate();
            XLSXFile.writeFile();
        }
        catch (Exception e) {
            System.err.println(e.getMessage());
            System.exit(1);
        }
    }
}