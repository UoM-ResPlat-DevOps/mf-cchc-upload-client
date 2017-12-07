import arc.mf.client.RemoteServer;
import arc.mf.client.ServerClient;
import arc.streams.StreamCopy;
import arc.xml.XmlDoc;
import arc.xml.XmlStringWriter;
import arc.xml.XmlDoc.Element;
import org.apache.commons.cli.*;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Iterator;
import java.util.List;

public class CchcUploader {

    public static void main(String[] args) {

//        CLI

        Options options = new Options();
        Option excelMetadataFileOption = Option.builder("excel")
                .required(true)
                .hasArg()
                .longOpt("excel-file")
                .desc("full path to the excel metadata file")
                .build();
        options.addOption(excelMetadataFileOption);

        Option hostOption = Option.builder("h")
                .required(true)
                .hasArg()
                .longOpt("host")
                .desc("mediaflux server hostname")
                .build();
        options.addOption(hostOption);

        Option portOption = Option.builder("p")
                .required(true)
                .hasArg()
                .longOpt("port")
                .desc("mediaflux server port number")
                .build();
        options.addOption(portOption);

        Option namespaceOption = Option.builder("ns")
                .required(true)
                .hasArg()
                .longOpt("namespace")
                .desc("mediaflux asset namespace")
                .build();
        options.addOption(namespaceOption);

        Option domainOption = Option.builder("d")
                .required(true)
                .hasArg()
                .longOpt("domain")
                .desc("mediaflux user domain")
                .build();
        options.addOption(domainOption);

        Option usernameOption = Option.builder("u")
                .required(false)
                .hasArg()
                .longOpt("username")
                .desc("mediaflux authentication username")
                .build();
        options.addOption(usernameOption);

        Option passwordOption = Option.builder("pw")
                .required(false)
                .hasArg()
                .longOpt("password")
                .desc("mediaflux authentication password")
                .build();
        options.addOption(passwordOption);

        Option dirPrefixOption = Option.builder("dp")
                .required(false)
                .hasArg()
                .longOpt("dirPrefix")
                .desc("directory prefix")
                .build();
        options.addOption(dirPrefixOption);

        Option rowStartOption = Option.builder("rs")
                .required(true)
                .hasArg()
//                .longOpt("rowStart")
                .desc("row start number")
                .build();
        options.addOption(rowStartOption);

        Option filmDocTypeOption = Option.builder("dt")
                .required(true)
                .hasArg()
                .longOpt("docType")
                .desc("mediaflux document type")
                .build();
        options.addOption(filmDocTypeOption);

        Option tokenAppOption = Option.builder("ta")
                .required(false)
                .hasArg()
                .longOpt("tokenApp")
                .desc("mediaflux authentication token app")
                .build();
        options.addOption(tokenAppOption);

        Option tokenOption = Option.builder("t")
                .required(false)
                .hasArg()
                .longOpt("token")
                .desc("mediaflux authentication token")
                .build();
        options.addOption(tokenOption);


        CommandLineParser parser = new DefaultParser();
        HelpFormatter formatter = new HelpFormatter();
        CommandLine cmd;
        String host, domain, username, password, excelFilePath, filmNameSpace, dirPrefix, filmDocType, tokenApp, token;
        int rowStart;
        int port = 80;
        boolean useTokenLogin = false;

        logToConsole("*********************************************************");
        logToConsole("CCHC Uploader");
        logToConsole("*********************************************************");

        try {
            cmd = parser.parse(options, args);

            domain = cmd.getOptionValue("domain");
            host = cmd.getOptionValue("h");
            port = Integer.parseInt(cmd.getOptionValue("port"));
            username = cmd.getOptionValue("username");
            password = cmd.getOptionValue("password");
            excelFilePath = cmd.getOptionValue("excel");
            filmNameSpace = cmd.getOptionValue("namespace");
            dirPrefix = cmd.getOptionValue("dp");
            rowStart = Integer.valueOf(cmd.getOptionValue("rs"));
            filmDocType = cmd.getOptionValue("docType");
            tokenApp = cmd.getOptionValue("tokenApp");
            token = cmd.getOptionValue("token");

            if (username == null || username.length() < 1) useTokenLogin = true;

            logToConsole("Command Line Arguments:");
            for (Option o : cmd.getOptions()) {
                if (o.hasLongOpt()){
                    if (o.getLongOpt().equalsIgnoreCase("password")
                            || o.getLongOpt().equalsIgnoreCase("username")) {
                        logToConsole(o.getLongOpt() + " = *********** ");
                    } else {
                        logToConsole(o.getLongOpt() + " = " + o.getValue());
                    }
                }
            }
            logToConsole("*********************************************************");

        } catch (ParseException e) {
            System.out.println(e.getMessage());
            formatter.printHelp("Film Uploader", options);
            System.exit(1);
            return;
        }

        try {
            Spreadsheet fs = new Spreadsheet(new FileInputStream(new File(excelFilePath)),
                                             rowStart,
                                             dirPrefix,
                                             filmDocType);
            ServerClient.setSessionPooling(false);
            boolean useHttps = false;
            if (port == 443) useHttps = true;
//            RemoteServer server = new RemoteServer(host, port, true, false);
//            RemoteServer server = new RemoteServer(host, port, true, true);
            RemoteServer server = new RemoteServer(host, port, true, useHttps);
            RemoteServer.Connection cxn = (RemoteServer.Connection) server.open();
            String sessionID = "";
            if (!useTokenLogin) {
                sessionID = cxn.connect(domain, username, password);
            } else {
                sessionID = cxn.connectWithToken(tokenApp, token);
            }
            uploadRecords(cxn, fs, filmNameSpace, dirPrefix, filmDocType);
        } catch (Throwable t) {
            t.printStackTrace();
        }
    }

    private static void uploadRecords(RemoteServer.Connection cxn,
                                    Spreadsheet fs,
                                    String nameSpace,
                                    String dirPrefix,
                                    String docType) throws  Throwable{
        int count = 1;
        for (Spreadsheet.Record rec : fs) {
            logToConsole("Attempt to upload : " + count + " of " + fs.length());
            uploadRecord(cxn, rec, nameSpace, dirPrefix, docType);
            count++;
        }

    }

    private static void uploadRecord(RemoteServer.Connection cxn,
                                      Spreadsheet.Record rec,
                                      String nameSpace,
                                      String dirPrefix,
                                      String docType) throws Throwable{

        Element e;
        List<Element> response;
        ServerClient.Input sci = null;
        String titleNs = rec.mediafluxFolder;
        String uploadNamespace = nameSpace + "/" + rec.mediafluxFolder;
        titleNs = rec.title.replaceAll("[^A-Za-z0-9 ()-]", "");
        createSubNameSpace(cxn, nameSpace, rec.mediafluxFolder);
        String fileLocation = null;
        logToConsole("[Upload Namespace] " + uploadNamespace);

        for (String slideFilename: rec.slideImages){
            fileLocation = FilenameUtils.concat(dirPrefix, slideFilename);
            logToConsole("[fileLocation] " + fileLocation);
            logToConsole("File in local path: " + String.valueOf(fileExistsLocal(fileLocation)));
            if (!fileExistsLocal(fileLocation)){
                logToConsole(" ** SKIP **. File missing in local path. Please check and try again: " + fileLocation);
            }
            else{
                sci = ServerClient.createInputFromURL("file:" + fileLocation);
                if(!assetExistsOnServer(cxn, rec, uploadNamespace, slideFilename, docType)){
                    e = cxn.execute("asset.create", rec.toXmlStringWriter(uploadNamespace, docType, slideFilename).document(), sci, null);
//                e = cxn.execute("asset.create", rec.toXmlStringWriter(uploadNamespace, docType, slideFilename).document());
                    response = e.elements();
                    logToConsole("[Upload Complete]  Title: " + rec.title + " >> Mediaflux Asset ID: " + response.get(0).toString());
                }
                else{
                    logToConsole("[File already exists on server]  Title: " + rec.title + ", Local File Path: " + fileLocation);
                }
            }
        }
        logToConsole("****");
    }

    public static boolean fileExistsLocal(String fullpath) throws Throwable{
        File f = new File(fullpath);
        if (!f.isFile() || !f.exists()){
            return false;
        }
        return true;
    }


    public static void logToConsole() {
        logToConsole("");
    }

    public static void logToConsole(String msg) {
        String timeStamp = new SimpleDateFormat("yyyy-MMM-dd HH:mm:ss").format(Calendar.getInstance().getTime());
        System.out.println("[" + timeStamp + "]  " + msg);
    }

    private static void createSubNameSpace(ServerClient.Connection cxn, String parent, String subNs) throws Throwable{
        XmlDoc.Element r;
        XmlStringWriter w = new XmlStringWriter();
        w.add("namespace", parent + "/" + subNs);
        r = cxn.execute("asset.namespace.exists", w.document());
        if (r.booleanValue("exists")) return;
        logToConsole(parent + "/" + subNs + ": " + String.valueOf(r.booleanValue("exists")));
        cxn.execute("asset.namespace.create", w.document());
    }

    private static boolean assetExistsOnServer(ServerClient.Connection cxn,
                                      Spreadsheet.Record rec,
                                      String namespace,
                                      String assetName,
                                      String filmDocType) throws Throwable{

        XmlStringWriter w = new XmlStringWriter();
        String q = "";
        String title = rec.title.replaceAll("'", "\\\\'");
        q += "name='" + assetName + "'";
        q += " and xpath( " + filmDocType + "/title)='" + title + "'";
        w.add("where", q);
        XmlDoc.Element r = cxn.execute("asset.query", w.document());
        if (r.count("id") > 0){
            logToConsole("[Film exists on server] ID: " + r.stringValue("id"));
            return true;
        }
        return false;
    }

}


