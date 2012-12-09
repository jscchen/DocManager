/*
 * Derived from the Apache POI example ModifyDocumentSummaryInformation.java from
 * the example source folder examples.src.org.apache.poi.hpsf.example
 */
package documentmanager;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;

import org.apache.poi.hpsf.CustomProperties;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.MarkUnsupportedException;
import org.apache.poi.hpsf.NoPropertySetStreamException;
import org.apache.poi.hpsf.PropertySet;
import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hpsf.UnexpectedPropertySetTypeException;
import org.apache.poi.hpsf.WritingNotSupportedException;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;


/**
 *This class is responsible for reading and writing the custom file properties
 *
 * @author Jessie Chen
 */
public class FilePropertyManager {
    private File fileSys;
    private POIFSFileSystem poifs;
    private DirectoryEntry fileDir;
    private SummaryInformation fileSumInfo;
    private DocumentSummaryInformation docSumInfo;
    private CustomProperties customProperties;

    private static final String TAG_PROPERTY_NAME = "Tag";
    private static final String PROCESS_PROPERTY_NAME_PREFIX = "processStep";
    private static final String PROCESS_PROPERTY_FLAG_SUFFIX = "_isCompleted";


    public FilePropertyManager(String fileName) throws IOException,
            NoPropertySetStreamException, MarkUnsupportedException,
            UnexpectedPropertySetTypeException
    {
        fileSys = new File(fileName);

        /* open the POI filesystem*/
        InputStream inStream = new FileInputStream(fileSys);
        poifs = new POIFSFileSystem(inStream);
        inStream.close();

        /* read the summary information*/
        fileDir = poifs.getRoot();
        try
        {
            DocumentEntry siEntry = (DocumentEntry) fileDir.getEntry(
                                    SummaryInformation.DEFAULT_STREAM_NAME);
            DocumentInputStream docInStream = new DocumentInputStream(siEntry);
            PropertySet ps = new PropertySet(docInStream);
            docInStream.close();
            fileSumInfo = new SummaryInformation(ps);
        }
        catch (FileNotFoundException ex)
        {
            /* no summary info, create one*/
            fileSumInfo = PropertySetFactory.newSummaryInformation();
        }

        /* read the document summary information */
        try
        {
            DocumentEntry dsiEntry = (DocumentEntry) fileDir.getEntry(
                                     DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            DocumentInputStream docInStream = new DocumentInputStream(dsiEntry);
            PropertySet ps = new PropertySet(docInStream);
            docInStream.close();
            docSumInfo = new DocumentSummaryInformation(ps);
        }
        catch (FileNotFoundException ex)
        {
            /* no document summary information yet, create one */
            docSumInfo = PropertySetFactory.newDocumentSummaryInformation();
        }

        //set the custom property handler
        customProperties = docSumInfo.getCustomProperties();
        if (customProperties == null){
            customProperties = new CustomProperties();
        }

    }

        public void setFileLastAuthor(String author) throws IOException, WritingNotSupportedException{
        /* Write the summary information and the document summary information
         * to the POI filesystem. */
        fileSumInfo.setLastAuthor(author);

        fileSumInfo.write(fileDir, SummaryInformation.DEFAULT_STREAM_NAME);

        /* Write the POI filesystem back to the original file. Please note that
         * in production code you should never write directly to the origin
         * file! In case of a writing error everything would be lost. */
        OutputStream out = new FileOutputStream(fileSys);
        poifs.writeFilesystem(out);
        out.close();
    }

    public void setFileComments(String comment) throws IOException, WritingNotSupportedException{
        /* Write the summary information and the document summary information
         * to the POI filesystem. */
        fileSumInfo.setComments(comment);

        fileSumInfo.write(fileDir, SummaryInformation.DEFAULT_STREAM_NAME);

        /* Write the POI filesystem back to the original file. Please note that
         * in production code you should never write directly to the origin
         * file! In case of a writing error everything would be lost. */
        OutputStream out = new FileOutputStream(fileSys);
        poifs.writeFilesystem(out);
        out.close();
    }

    public void setFileTags(String tags) throws IOException, WritingNotSupportedException{
        
        //check if there is custom property for tags
        if(customProperties.containsKey(TAG_PROPERTY_NAME)){
            customProperties.remove(TAG_PROPERTY_NAME);
        }

        customProperties.put(TAG_PROPERTY_NAME, tags);
        //System.out.println("New Tags: " + customProperties.get(TAG_PROPERTY_NAME));
        /* Write the custom properties back to the document summary
         * information. */
        docSumInfo.setCustomProperties(customProperties);

        /* Write the summary information and the document summary information
         * to the POI filesystem. */
        fileSumInfo.write(fileDir, SummaryInformation.DEFAULT_STREAM_NAME);
        docSumInfo.write(fileDir, DocumentSummaryInformation.DEFAULT_STREAM_NAME);

        /* Write the POI filesystem back to the original file. Please note that
         * in production code you should never write directly to the origin
         * file! In case of a writing error everything would be lost. */
        OutputStream out = new FileOutputStream(fileSys);
        poifs.writeFilesystem(out);
        out.close();
    }

    public void addProcessStep(String stepNum, String stepDescription) throws IOException, WritingNotSupportedException{
         //check if there is custom property for tags
        String propertyKey = PROCESS_PROPERTY_NAME_PREFIX + stepNum;
        String processStepKey = propertyKey + PROCESS_PROPERTY_FLAG_SUFFIX;

        if(customProperties.containsKey(propertyKey)){
            //TO-DO: decide whether to overwrite or give error
            customProperties.remove(propertyKey);
        }

        customProperties.put(propertyKey, stepDescription);
        //this property is for tracking if the process step is completed
        customProperties.put(processStepKey, false);
        
        //System.out.println("New Tags: " + customProperties.get(TAG_PROPERTY_NAME));
        /* Write the custom properties back to the document summary
         * information. */
        docSumInfo.setCustomProperties(customProperties);

        /* Write the summary information and the document summary information
         * to the POI filesystem. */
        fileSumInfo.write(fileDir, SummaryInformation.DEFAULT_STREAM_NAME);
        docSumInfo.write(fileDir, DocumentSummaryInformation.DEFAULT_STREAM_NAME);

        /* Write the POI filesystem back to the original file. Please note that
         * in production code you should never write directly to the origin
         * file! In case of a writing error everything would be lost. */
        OutputStream out = new FileOutputStream(fileSys);
        poifs.writeFilesystem(out);
        out.close();
    }

    public void editProcessInfo(String stepNum, String stepDescription, String status)
            throws IOException, WritingNotSupportedException{
        //check if there is custom property for tags
        String propertyKey = PROCESS_PROPERTY_NAME_PREFIX + stepNum;
        String processStepKey = propertyKey + PROCESS_PROPERTY_FLAG_SUFFIX;
        boolean statusFlag = false;
        if(status.equals("Yes")){
            statusFlag = true;
        }

        System.out.println(stepNum + " Status: " + statusFlag);

        if(customProperties.containsKey(propertyKey)){
            //TO-DO: decide whether to overwrite or give error
            customProperties.remove(propertyKey);
        }
       customProperties.put(propertyKey, stepDescription);
       
       //update the process step status flag
       if(customProperties.containsKey(processStepKey)){
            //TO-DO: decide whether to overwrite or give error
            customProperties.remove(processStepKey);
        }
        customProperties.put(processStepKey, statusFlag);
        
        //System.out.println("New Tags: " + customProperties.get(TAG_PROPERTY_NAME));
        /* Write the custom properties back to the document summary
         * information. */
        docSumInfo.setCustomProperties(customProperties);

        /* Write the summary information and the document summary information
         * to the POI filesystem. */
        fileSumInfo.write(fileDir, SummaryInformation.DEFAULT_STREAM_NAME);
        docSumInfo.write(fileDir, DocumentSummaryInformation.DEFAULT_STREAM_NAME);

        /* Write the POI filesystem back to the original file. Please note that
         * in production code you should never write directly to the origin
         * file! In case of a writing error everything would be lost. */
        OutputStream out = new FileOutputStream(fileSys);
        poifs.writeFilesystem(out);
        out.close();
    }


    public String getFileAuthor(){
        String authorName = fileSumInfo.getAuthor();

        return authorName;
    }

    public String getFileLastAuthor(){
        String authorName = fileSumInfo.getLastAuthor();

        return authorName;
    }

    public String getFileComments(){
        String fileComments = fileSumInfo.getComments();

        return fileComments;
    }

    public String getFileTags(){
        Object tagValue = null;
        if(customProperties.containsKey(TAG_PROPERTY_NAME)){
            tagValue = customProperties.get(TAG_PROPERTY_NAME);
        }else{
            tagValue = "";
        }

        return (String) tagValue;
    }

    public HashMap <String, String> getFileProcessStepDescription(){
        HashMap stepsDescriptionMap = new HashMap();
        Object [] stepKeys = customProperties.keySet().toArray();
        int keyCount = stepKeys.length;
        int count = 0;
        int keyNumIndex = PROCESS_PROPERTY_NAME_PREFIX.length();
        for(count=0; count < keyCount; count++){
            //only get the custom property that are process related
            String key = stepKeys[count].toString();
            if(key.startsWith(PROCESS_PROPERTY_NAME_PREFIX) &&
                    !(key.endsWith(PROCESS_PROPERTY_FLAG_SUFFIX))){
                String keyNum = key.substring(keyNumIndex);
                String description = customProperties.get(key).toString();

                stepsDescriptionMap.put(keyNum,description);
            }
        }

        return stepsDescriptionMap;
    }

    public HashMap <String, String> getFileProcessStepStatus(){
        HashMap stepsStatusMap = new HashMap();
        Object [] stepKeys = customProperties.keySet().toArray();
        int keyCount = stepKeys.length;
        int count = 0;
        int keyNumIndex = PROCESS_PROPERTY_NAME_PREFIX.length();
        for(count=0; count < keyCount; count++){
            //only get the custom property that are process related
            String key = stepKeys[count].toString();
            if(key.startsWith(PROCESS_PROPERTY_NAME_PREFIX) &&
                    key.endsWith(PROCESS_PROPERTY_FLAG_SUFFIX)){
                String keyNum = key.substring(keyNumIndex, keyNumIndex+1);
                String status = customProperties.get(key).toString();
                if(status.equals("false")){
                    status = "No";
                }else{
                    status = "Yes";
                }

                stepsStatusMap.put(keyNum,status);
            }
        }

        return stepsStatusMap;
    }
}
