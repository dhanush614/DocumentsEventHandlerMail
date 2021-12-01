package sample.actionhandler;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.InternetHeaders;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.util.ByteArrayDataSource;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.filenet.api.collection.ContentElementList;
import com.filenet.api.collection.FolderSet;
import com.filenet.api.constants.AutoClassify;
import com.filenet.api.constants.AutoUniqueName;
import com.filenet.api.constants.CheckinType;
import com.filenet.api.constants.DefineSecurityParentage;
import com.filenet.api.constants.PropertyNames;
import com.filenet.api.constants.RefreshMode;
import com.filenet.api.core.ContentTransfer;
import com.filenet.api.core.Document;
import com.filenet.api.core.Factory;
import com.filenet.api.core.Folder;
import com.filenet.api.core.ObjectStore;
import com.filenet.api.core.ReferentialContainmentRelationship;
import com.filenet.api.engine.EventActionHandler;
import com.filenet.api.events.ObjectChangeEvent;
import com.filenet.api.property.FilterElement;
import com.filenet.api.property.Properties;
import com.filenet.api.property.PropertyFilter;
import com.filenet.api.util.Id;
import com.ibm.casemgmt.api.Case;
import com.ibm.casemgmt.api.CaseType;
import com.ibm.casemgmt.api.context.CaseMgmtContext;
import com.ibm.casemgmt.api.context.P8ConnectionCache;
import com.ibm.casemgmt.api.context.SimpleP8ConnectionCache;
import com.ibm.casemgmt.api.context.SimpleVWSessionCache;
import com.ibm.casemgmt.api.objectref.ObjectStoreReference;
import com.ibm.casemgmt.api.properties.CaseMgmtProperties;

public class DocumentsEventHandler implements EventActionHandler {
	public void onEvent(ObjectChangeEvent event, Id subId) {
		System.out.println("Inside onEvent method");
		CaseMgmtContext origCmctx = null;
		try {
			int caseCount = 0;
			P8ConnectionCache connCache = new SimpleP8ConnectionCache();
			origCmctx = CaseMgmtContext.set(new CaseMgmtContext(new SimpleVWSessionCache(), connCache));
			ObjectStore os = event.getObjectStore();
			System.out.println("OS" + os);
			ObjectStoreReference targetOsRef = new ObjectStoreReference(os);
			System.out.println("TOS" + targetOsRef);
			Id id = event.get_SourceObjectId();
			FilterElement fe = new FilterElement(null, null, null, "Owner Name", null);
			PropertyFilter pf = new PropertyFilter();
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.CONTENT_SIZE, null));
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.CONTENT_ELEMENTS, null));
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.FOLDERS_FILED_IN, null));
			pf.addIncludeProperty(fe);
			Document doc = Factory.Document.fetchInstance(os, id, pf);
			System.out.println("Document Name" + doc.get_Name());
			ContentElementList docContentList = doc.get_ContentElements();
			Iterator iter = docContentList.iterator();
			while (iter.hasNext()) {
				ContentTransfer ct = (ContentTransfer) iter.next();
				InputStream stream = ct.accessContentStream();
				int rowLastCell = 0;
				HashMap<Integer, String> headers = new HashMap<Integer, String>();
				HashMap<String, String> propDescMap = new HashMap<String, String>();
				XSSFWorkbook workbook = new XSSFWorkbook(stream);
				XSSFSheet sheet = workbook.getSheetAt(0);
				XSSFSheet sheet1 = workbook.getSheetAt(1);
				Iterator<Row> rowIterator = sheet.iterator();
				Iterator<Row> rowIterator1 = sheet1.iterator();
				while (rowIterator1.hasNext()) {
					Row row = rowIterator1.next();
					if (row.getRowNum() > 0) {
						String key = null, value = null;
						key = row.getCell(0).getStringCellValue();
						value = row.getCell(1).getStringCellValue();
						if (key != null && value != null) {
							propDescMap.put(key, value);
						}
					}
				}
				String headerValue;
				if (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					Iterator<Cell> cellIterator = row.cellIterator();
					int colNum = 0;
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						headerValue = cell.getStringCellValue();
						if (headerValue.contains("*")) {
							if (headerValue.contains("datetime")) {
								headerValue = headerValue.replaceAll("\\* *\\([^)]*\\) *", "").trim();
								headerValue += "dateField";
							} else {
								headerValue = headerValue.replaceAll("\\* *\\([^)]*\\) *", "").trim();
							}
						}
						if (headerValue.contains("datetime")) {
							headerValue = headerValue.replaceAll("\\([^)]*\\) *", "").trim();
							headerValue += "dateField";
						} else {
							headerValue = headerValue.replaceAll("\\([^)]*\\) *", "").trim();
						}
						headers.put(colNum++, headerValue);
					}
					rowLastCell = row.getLastCellNum();
					Cell cell1 = row.createCell(rowLastCell, Cell.CELL_TYPE_STRING);
					if (row.getRowNum() == 0) {
						cell1.setCellValue("Status");
					}

				}
				CaseType caseType = CaseType.fetchInstance(targetOsRef, doc.get_Name());
				int rowStart = sheet.getFirstRowNum() + 1;
				int rowEnd = sheet.getLastRowNum();
				for (int rowNumber = rowStart; rowNumber <= rowEnd; rowNumber++) {
					Row row = sheet.getRow(rowNumber);
					if (row == null) {
						break;
					} else {
						int colNum = 0;
						String caseId = "";
						try {
							Case pendingCase = Case.createPendingInstance(caseType);
							CaseMgmtProperties caseMgmtProperties = pendingCase.getProperties();
							for (int i = 0; i < row.getLastCellNum(); i++) {
								Cell cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
								try {
									if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
										colNum++;
									} else {
										if (headers.get(colNum).contains("dateField")) {
											if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
													&& DateUtil.isCellDateFormatted(cell)) {
												String symName = headers.get(colNum).replace("dateField", "");
												Date date = cell.getDateCellValue();
												caseMgmtProperties.putObjectValue(propDescMap.get(symName), date);
												colNum++;
											} else {
												colNum++;
											}
										} else {
											caseMgmtProperties.putObjectValue(propDescMap.get(headers.get(colNum++)),
													getCharValue(cell));
										}
									}
								} catch (Exception e) {
									System.out.println(e);
									e.printStackTrace();
								}
							}
							System.out.println("Case Creation");
							pendingCase.save(RefreshMode.REFRESH, null, null);
							caseId = pendingCase.getId().toString();
							System.out.println("Case_ID: " + caseId);
							Cell cell1 = row.createCell(rowLastCell);
							if (!caseId.isEmpty()) {
								caseCount += 1;
								System.out.println("CaseCount: " + caseCount);
								cell1.setCellValue("Success");
							} else {
								cell1.setCellValue("Failure");
							}
						} catch (Exception e) {
							System.out.println(e);
							e.printStackTrace();
						}
					}
				}
				InputStream is = null;
				ByteArrayOutputStream bos = null;
				try {
					bos = new ByteArrayOutputStream();
					workbook.write(bos);
					byte[] barray = bos.toByteArray();
					is = new ByteArrayInputStream(barray);
					String docTitle = doc.get_Name();
					FolderSet folderSet = doc.get_FoldersFiledIn();
					Folder folder = null;
					Iterator<Folder> folderSetIterator = folderSet.iterator();
					if (folderSetIterator.hasNext()) {
						folder = folderSetIterator.next();
					}
					String folderPath = folder.get_PathName();
					folderPath += " Response";
					Folder responseFolder = Factory.Folder.fetchInstance(os, folderPath, null);
					updateDocument(os, is, doc, docTitle, responseFolder);
					sendEmail(barray);
				} catch (Exception e) {
					System.out.println(e);
					e.printStackTrace();
				} finally {
					if (bos != null) {
						bos.close();
					}
					if (is != null) {
						is.close();
					}
					if (stream != null) {
						stream.close();
					}
				}
			}
		} catch (Exception e) {
			System.out.println(e);
			e.printStackTrace();
			throw new RuntimeException(e);
		} finally {
			CaseMgmtContext.set(origCmctx);
		}
	}

	private void updateDocument(ObjectStore os, InputStream is, Document doc, String docTitle, Folder responseFolder) {
		// TODO Auto-generated method stub
		String docClassName = doc.getClassName() + "Response";
		Document updateDoc = Factory.Document.createInstance(os, docClassName);
		ContentElementList contentList = Factory.ContentElement.createList();
		ContentTransfer contentTransfer = Factory.ContentTransfer.createInstance();
		contentTransfer.setCaptureSource(is);
		contentTransfer.set_RetrievalName(docTitle + ".xlsx");
		contentTransfer.set_ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		contentList.add(contentTransfer);

		updateDoc.set_ContentElements(contentList);
		updateDoc.checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MAJOR_VERSION);
		Properties p = updateDoc.getProperties();
		p.putValue("DocumentTitle", docTitle);
		updateDoc.setUpdateSequenceNumber(null);
		updateDoc.save(RefreshMode.REFRESH);
		ReferentialContainmentRelationship rc = responseFolder.file(updateDoc, AutoUniqueName.AUTO_UNIQUE, docTitle,
				DefineSecurityParentage.DO_NOT_DEFINE_SECURITY_PARENTAGE);
		rc.save(RefreshMode.REFRESH);
	}

	private static Object getCharValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			return cell.getNumericCellValue();

		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		}
		return null;
	}

	private void sendEmail(byte[] barray) throws AddressException, MessagingException, IOException {

		Session session = Session.getDefaultInstance(new java.util.Properties());

		MimeMessage message = new MimeMessage(session);
		message.setFrom(new InternetAddress("default@ibmdba.com"));
		message.addRecipient(Message.RecipientType.TO, new InternetAddress("jukka@ibmdba.com"));
		message.setSubject("Bulk Case Creation Response Sheet");

		BodyPart messageBodyPart1 = new MimeBodyPart();
		messageBodyPart1
				.setText("Hi,\n\nThis e-mail is to notify you that the request for Bulk Case Creation is completed."
						+ "\n\nPlease refer the attachment for Case Creation Status\n\nRegards,\nBulk Case Creation Team");

		MimeBodyPart messageBodyPart2 = new MimeBodyPart();
		String filename = "LA_LoanProcessingCaseTypeResponse.xlsx";
		DataSource attachment = new ByteArrayDataSource(barray, "application/vnd.ms-excel");
		messageBodyPart2.setDataHandler(new DataHandler(attachment));
		messageBodyPart2.setFileName(filename);

		Multipart multipart = new MimeMultipart();
		multipart.addBodyPart(messageBodyPart1);
		multipart.addBodyPart(messageBodyPart2);

		message.setContent(multipart);

		Transport.send(message);

		System.out.println("Mail sent....");

	}
}
