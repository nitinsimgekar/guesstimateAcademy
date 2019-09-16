package com.guesstimate.testlink;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

public class XmlToXls {

	private static String[] columns = { "Test Case Title", "Summary", "Preconditions", "Test Case Steps", "Expected",
			"Priority", "Status", "Review Comments", "Need Update In TestLink", "Status", "buildName", "Notes",
			"BugID" };

	public static void main(String[] args)
			throws ParserConfigurationException, SAXException, IOException, InvalidFormatException {

		File file = new File(args[0]);

		Workbook workbook = new XSSFWorkbook();

		XmlToXls xmlToCSV = new XmlToXls();
		xmlToCSV.readXML(file, workbook);

		FileOutputStream fileOut = new FileOutputStream("parsed-output.xls");
		workbook.write(fileOut);
		fileOut.close();

		workbook.close();
	}

	private void readTestSuite(Workbook workbook, NodeList testsuite, String suiteTitle2) {

		for (int k = 0; k < testsuite.getLength(); k++) {

			Node testsuiteNode = testsuite.item(k);

			if (testsuiteNode.getNodeType() == Node.ELEMENT_NODE) {

				Element testsuiteElement = (Element) testsuiteNode;

				NodeList innerTestSuite = testsuiteElement.getElementsByTagName("testsuite");

				String suiteTitle = testsuiteElement.getAttribute("name");

				if (!suiteTitle.isEmpty()) {

					String text = "";

					if (!suiteTitle2.isEmpty())
						text = suiteTitle2 + "-" + suiteTitle;
					else
						text = suiteTitle;

					System.out.println("Parsing test cases for: " + text);

					try {
						Sheet sheet = workbook.createSheet(text);

						Row headerRow = sheet.createRow(0);

						Font headerFont = workbook.createFont();
						headerFont.setBold(true);
						headerFont.setFontHeightInPoints((short) 14);
						headerFont.setColor(IndexedColors.RED.getIndex());

						CellStyle headerCellStyle = workbook.createCellStyle();
						headerCellStyle.setFont(headerFont);

						for (int i = 0; i < columns.length; i++) {
							Cell cell = headerRow.createCell(i);
							cell.setCellValue(columns[i]);
							cell.setCellStyle(headerCellStyle);
						}

						readTestCase(testsuiteElement, sheet);

						if (innerTestSuite.getLength() > 0) {

							readTestSuite(workbook, innerTestSuite, suiteTitle);
						}
					} catch (IllegalArgumentException e) {

					}

				}
			}
		}

	}

	private void readTestCase(Element testsuiteElement, Sheet sheet) {

		NodeList nList = testsuiteElement.getElementsByTagName("testcase");

		int rowNum = 1;

		for (int temp = 0; temp < nList.getLength(); temp++) {

			Node nNode = nList.item(temp);

			if (nNode.getNodeType() == Node.ELEMENT_NODE) {

				Row row = sheet.createRow(rowNum++);

				Element eElement = (Element) nNode;

				String title = eElement.getAttribute("name");
				String summary = eElement.getElementsByTagName("summary").item(0).getTextContent();
				String preconditions = eElement.getElementsByTagName("preconditions").item(0).getTextContent();

				row.createCell(0).setCellValue(title);
				row.createCell(1).setCellValue(summary);
				row.createCell(2).setCellValue(preconditions);

				StringBuilder sbactions = new StringBuilder();
				StringBuilder sbexpectedresults = new StringBuilder();

				NodeList steps = eElement.getElementsByTagName("steps");

				for (int i = 0; i < steps.getLength(); i++) {

					Node stepsNode = steps.item(i);

					if (stepsNode.getNodeType() == Node.ELEMENT_NODE) {

						Element stepsElement = (Element) stepsNode;

						NodeList stepsNodeList = stepsElement.getElementsByTagName("step");

						for (int j = 0; j < stepsNodeList.getLength(); j++) {

							Node stepNode = stepsNodeList.item(j);

							if (stepNode.getNodeType() == Node.ELEMENT_NODE) {

								Element stepElement = (Element) stepNode;

								String actions = stepElement.getElementsByTagName("actions").item(0).getTextContent();
								sbactions.append(actions + "\n");
								String expectedresults = stepElement.getElementsByTagName("expectedresults").item(0)
										.getTextContent();
								sbexpectedresults.append(expectedresults + "\n");

							}

						}

					}

				}

				row.createCell(3).setCellValue(sbactions.toString());
				row.createCell(4).setCellValue(sbexpectedresults.toString());
				row.createCell(5).setCellValue("Medium");
				row.createCell(6).setCellValue("Final");
				// row.createCell(8).setCellValue("ExecutionResult");
			}

		}

		for (int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);
		}

	}

	private void readXML(File file, Workbook workbook) throws ParserConfigurationException, SAXException, IOException {

		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse(file);

		doc.getDocumentElement().normalize();

		Element testsuites = doc.getDocumentElement();

		NodeList testsuite = testsuites.getChildNodes();

		readTestSuite(workbook, testsuite, "");

	}

}
