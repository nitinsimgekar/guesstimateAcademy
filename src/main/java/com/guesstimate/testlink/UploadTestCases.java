package com.guesstimate.testlink;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.ini4j.Ini;

import br.eti.kinoshita.testlinkjavaapi.TestLinkAPI;
import br.eti.kinoshita.testlinkjavaapi.constants.ActionOnDuplicate;
import br.eti.kinoshita.testlinkjavaapi.constants.ExecutionStatus;
import br.eti.kinoshita.testlinkjavaapi.constants.ExecutionType;
import br.eti.kinoshita.testlinkjavaapi.constants.TestCaseStatus;
import br.eti.kinoshita.testlinkjavaapi.constants.TestImportance;
import br.eti.kinoshita.testlinkjavaapi.model.Build;
import br.eti.kinoshita.testlinkjavaapi.model.TestCase;
import br.eti.kinoshita.testlinkjavaapi.model.TestCaseStep;
import br.eti.kinoshita.testlinkjavaapi.model.TestPlan;
import br.eti.kinoshita.testlinkjavaapi.model.TestProject;
import br.eti.kinoshita.testlinkjavaapi.model.TestSuite;
import br.eti.kinoshita.testlinkjavaapi.util.TestLinkAPIException;

public class UploadTestCases {

	/*
	 * the data is stored in ini file follows the below format. tc id + tc VersionId
	 * + "|false/true (isCustomfield mapped)"
	 */

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {

		UploadTestCases uploadTestCases = new UploadTestCases();
		TestLinkAPI api = getInstance(args[0],args[1]);

		System.out.println("Sit back with a irani chai until system upload testcases to testlink.");

		uploadTestCases.readxls(args[0], api);

	}

	public static TestLinkAPI getInstance(String url, String devKey) {

		TestLinkAPI api = null;

		URL testlinkURL = null;

		try {
			testlinkURL = new URL(url);
		} catch (MalformedURLException mue) {
			mue.printStackTrace(System.err);
			System.exit(-1);
		}

		try {
			api = new TestLinkAPI(testlinkURL, devKey);
		} catch (TestLinkAPIException te) {
			te.printStackTrace(System.err);
			System.exit(-1);
		}

		System.out.println(api.ping());

		return api;
	}

	private void readxls(String pathName, TestLinkAPI api)
			throws EncryptedDocumentException, InvalidFormatException, IOException {

		File file = new File(pathName);
		FileInputStream excelFile = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(excelFile);

		TestProject project = getTestLinkProject(api, file.getName().substring(0, file.getName().lastIndexOf(".")));

		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

			Sheet sheet = workbook.getSheetAt(i);

			String sheetName = sheet.getSheetName().trim();

			TestPlan testPlan = getTestPlan(api, project);

			String[] sheetNameList = sheetName.split("-");

			TestSuite suite = getTestSuite(api, sheetNameList[0].trim(), project.getId());

			Integer suiteId = suite.getId();

			for (int m = 1; m < sheetNameList.length; m++) {
				suite = getSubTestSuite(api, sheetNameList[m].trim(), suiteId, project.getId());
				suiteId = suite.getId();
			}

			DataFormatter dataFormatter = new DataFormatter();

			int k = 0;
			int l = 0;

			boolean isNewTestCasesAdded = false;

			List<String> uploadFailed = new ArrayList<String>();
			List<String> looksNewTestCase = new ArrayList<String>();
			List<String> setExecutionFailed = new ArrayList<String>();
			List<String> possibleDuplicates = new ArrayList<String>();
			List<Integer> emptyTitle = new ArrayList<Integer>();
			List<Integer> invalidTitle = new ArrayList<Integer>();

			for (int j = 1; j < sheet.getPhysicalNumberOfRows(); j++) {

				Row row = sheet.getRow(j);

				if (!dataFormatter.formatCellValue(row.getCell(0)).trim().isEmpty()) {

					if (dataFormatter.formatCellValue(row.getCell(0)).contains(" ")
							|| dataFormatter.formatCellValue(row.getCell(0)).substring(0, 1).matches("[a-z]")) {

						if (dataFormatter.formatCellValue(row.getCell(6)).equalsIgnoreCase("final")) {

							if (testCaseGenerated(project.getName(), sheetName,
									dataFormatter.formatCellValue(row.getCell(0))) == null) {

								TestCase tc = createTestCase(api, row, dataFormatter, project.getId(), suite.getId());

								if (tc != null) {
									isNewTestCasesAdded = true;
									setTestCaseGenerated(project.getName(), sheetName,
											dataFormatter.formatCellValue(row.getCell(0)),
											tc.getId() + "," + tc.getVersion() + ",false");
									k++;
								} else
									uploadFailed.add(dataFormatter.formatCellValue(row.getCell(0)));
							} else
								possibleDuplicates.add(dataFormatter.formatCellValue(row.getCell(0)));
						}

						if (!isNewTestCasesAdded) {

							try {

								String[] tcid = testCaseGenerated(project.getName(), sheetName,
										dataFormatter.formatCellValue(row.getCell(0))).split(",");

								if (dataFormatter.formatCellValue(row.getCell(8)).equalsIgnoreCase("TestCaseUpdate")) {

									possibleDuplicates.remove(dataFormatter.formatCellValue(row.getCell(0)));

									if (updateTestCase(api, row, dataFormatter, project.getId(), suite.getId(),
											Integer.valueOf(tcid[0]))) {

										l++;

									} else
										uploadFailed.add(dataFormatter.formatCellValue(row.getCell(0)));

								} else if (dataFormatter.formatCellValue(row.getCell(8))
										.equalsIgnoreCase("ExecutionResult")) {

									possibleDuplicates.remove(dataFormatter.formatCellValue(row.getCell(0)));

									Build[] builds = api.getBuildsForTestPlan(testPlan.getId());

									boolean isbuildFound = false;

									for (Build build : builds) {

										if (build.getName()
												.equalsIgnoreCase(dataFormatter.formatCellValue(row.getCell(10)))) {

											if (updateExecutionReport(api, row, dataFormatter, project.getId(),
													suite.getId(), Integer.valueOf(tcid[0]), testPlan.getId(), build)) {
												isbuildFound = true;
												l++;
											}

										}
									}

									if (!isbuildFound)
										setExecutionFailed.add(dataFormatter.formatCellValue(row.getCell(0)));
								}

							} catch (NullPointerException e) {
								if (!uploadFailed.contains(dataFormatter.formatCellValue(row.getCell(0))))
									looksNewTestCase.add(dataFormatter.formatCellValue(row.getCell(0)));
							}

						}
					} else
						invalidTitle.add(row.getRowNum());
				} else
					emptyTitle.add(row.getRowNum());

			}

			if (k != 0) {
				System.out.println("uploaded " + k + " testcases out of " + (sheet.getPhysicalNumberOfRows() - 1));
			}

			if (l != 0) {
				System.out.println("Updated " + l + " testcases out of " + (sheet.getPhysicalNumberOfRows() - 1));
			}

			if (uploadFailed.size() != 0) {
				System.out.println("Test steps and expected results donot match properly for TC: " + uploadFailed);
			}

			if (looksNewTestCase.size() != 0) {
				System.out
						.println("This looks like new test case, which you are trying to update: " + looksNewTestCase);
			}

			if (setExecutionFailed.size() != 0) {
				System.out.println("Build name doesnt match: " + setExecutionFailed);
			}

			if (invalidTitle.size() != 0) {
				System.out.println("Rows with invalid title: " + invalidTitle);
			}

			if (possibleDuplicates.size() != 0 && k != 0) {
				System.out.println("Possible duplicate titles: " + possibleDuplicates);
			}
			if (emptyTitle.size() != 0) {
				System.out.println("Rows with empty title: " + emptyTitle);
			}

			System.out.println("=================");
		}

		workbook.close();

	}

	private String testCaseGenerated(String testProject, String testSuite, String title) {

		String id = null;
		try {
			String iniFileName = "./res/testcasexls/" + testProject.trim() + ".ini";
			File file = new File(iniFileName);

			if (!file.exists()) {
				file.createNewFile();
			}

			Ini ini = new Ini(new FileReader(file));
			id = ini.get(testSuite.trim(), title.trim());

		} catch (IOException e) {
			e.printStackTrace();
		}

		return id;
	}

	public void setTestCaseGenerated(String testProject, String testSuite, String title, String string) {

		try {
			String iniFileName = "./res/testcasexls/" + testProject.trim() + ".ini";
			Ini ini = new Ini(new FileReader(new File(iniFileName)));

			ini.put(testSuite.trim(), title.trim(), string.trim());
			ini.store(new File(iniFileName));

		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private TestCase createTestCase(TestLinkAPI api, Row row, DataFormatter dataFormatter, Integer projectId,
			Integer suiteId) {

		String[] stepsLines = dataFormatter.formatCellValue(row.getCell(3)).trim().split("\n");
		String[] expectedLines = dataFormatter.formatCellValue(row.getCell(4)).trim().split("\n");

		if (expectedLines.length < stepsLines.length) {
			return null;
		}

		if (expectedLines.length > stepsLines.length) {

			int k = stepsLines.length - 1;
			StringBuilder builder = new StringBuilder();
			for (int i = k; i < expectedLines.length; i++) {
				builder.append(expectedLines[i] + "\n");
			}

			expectedLines[k] = builder.toString();

		}

		List<TestCaseStep> steps = new ArrayList<TestCaseStep>();

		for (int i = 0; i < stepsLines.length; i++) {

			TestCaseStep step = new TestCaseStep();
			step.setNumber(i + 1);
			step.setExpectedResults(expectedLines[i]);
			step.setExecutionType(ExecutionType.MANUAL);
			step.setActions(stepsLines[i]);
			steps.add(step);
		}

		TestImportance importance = null;

		try {
			importance = TestImportance.valueOf(dataFormatter.formatCellValue(row.getCell(5)).toUpperCase());
		} catch (IllegalArgumentException e) {
			importance = TestImportance.MEDIUM;
		}

		TestCase tc = null;
		try {

			tc = api.createTestCase(dataFormatter.formatCellValue(row.getCell(0)), suiteId, projectId, "smiriyala",
					dataFormatter.formatCellValue(row.getCell(1)), steps, dataFormatter.formatCellValue(row.getCell(2)),
					TestCaseStatus.FINAL, importance, ExecutionType.MANUAL, null, null, true,
					ActionOnDuplicate.GENERATE_NEW);

		} catch (TestLinkAPIException e) {

		}

		return tc;

	}

	private boolean updateExecutionReport(TestLinkAPI api, Row row, DataFormatter dataFormatter, Integer projectId,
			Integer suiteId, Integer tcid, Integer testPlanId, Build build) {

		ExecutionStatus executionStatus = ExecutionStatus
				.valueOf(dataFormatter.formatCellValue(row.getCell(9)).toUpperCase());

		try {
			api.setTestCaseExecutionResult(tcid, null, testPlanId, executionStatus, build.getId(), build.getName(),
					dataFormatter.formatCellValue(row.getCell(11)), null,
					dataFormatter.formatCellValue(row.getCell(12)), null, null, null, null);
		} catch (TestLinkAPIException e) {
			return false;
		}

		return true;

	}

	private TestProject getTestLinkProject(TestLinkAPI api, String testProjectName) {

		TestProject project = null;

		try {

			project = api.getTestProjectByName(testProjectName);

		} catch (TestLinkAPIException e) {

			String[] notes = testProjectName.split("(?<!(^|[A-Z]))(?=[A-Z])|(?<!^)(?=[A-Z][a-z])");

			StringBuilder note = new StringBuilder();
			StringBuilder prefex = new StringBuilder();

			for (String string : notes) {
				note.append(string + " ");
				prefex.append(string.substring(0, 1));
			}

			project = api.createTestProject(testProjectName, prefex.toString().toUpperCase(),
					note.toString().toLowerCase(), true, true, true, false, true, true);

		}

		System.out.println("Test project: " + testProjectName);

		return project;

	}

	private TestSuite getTestSuite(TestLinkAPI api, String testSuiteName, Integer testProjectId) {

		TestSuite testSuite = null;

		boolean isSuiteExits = false;

		try {

			TestSuite[] testSuites = api.getFirstLevelTestSuitesForTestProject(testProjectId);

			for (TestSuite suite : testSuites) {
				if (suite.getName().equalsIgnoreCase(testSuiteName)) {
					testSuite = suite;
					isSuiteExits = true;
					break;
				}
			}

			if (!isSuiteExits) {

				testSuite = api.createTestSuite(testProjectId, testSuiteName, testSuiteName, null, null, true,
						ActionOnDuplicate.GENERATE_NEW);

			}

			System.out.println("Test suite: " + testSuite.getName());

		} catch (TestLinkAPIException e) {
			e.printStackTrace(System.err);
			System.exit(-1);
		}

		return testSuite;

	}

	private TestSuite getSubTestSuite(TestLinkAPI api, String testSuiteName, Integer testSuiteId,
			Integer testProjectId) {

		TestSuite testSuite = null;

		boolean isSuiteExits = false;

		try {

			TestSuite[] testSuites = api.getTestSuitesForTestSuite(testSuiteId);

			for (TestSuite suite : testSuites) {
				if (suite.getName().equalsIgnoreCase(testSuiteName)) {
					testSuite = suite;
					isSuiteExits = true;
					break;
				}
			}

			if (!isSuiteExits) {

				testSuite = api.createTestSuite(testProjectId, testSuiteName, testSuiteName, testSuiteId, null, true,
						ActionOnDuplicate.GENERATE_NEW);

			}

			System.out.println("Test suite: " + testSuite.getName());

		} catch (TestLinkAPIException e) {
			e.printStackTrace(System.err);
			System.exit(-1);
		}

		return testSuite;

	}

	private TestPlan getTestPlan(TestLinkAPI api, TestProject project) {

		TestPlan testPlan = null;

		try {

			TestPlan[] plans = api.getProjectTestPlans(project.getId());

			String[] notes = project.getName().split("(?<!(^|[A-Z]))(?=[A-Z])|(?<!^)(?=[A-Z][a-z])");

			StringBuilder note = new StringBuilder();

			for (String string : notes) {
				note.append(string + " ");
			}

			if (plans.length == 0) {
				testPlan = api.createTestPlan(project.getName(), project.getName(), note.toString().toLowerCase(), true,
						true);

			} else {

				testPlan = plans[0];

			}

			System.out.println("Test plan: " + testPlan.getName());

		} catch (TestLinkAPIException e) {
			e.printStackTrace(System.err);
			System.exit(-1);
		}

		return testPlan;

	}

	private boolean updateTestCase(TestLinkAPI api, Row row, DataFormatter dataFormatter, Integer projectId,
			Integer suiteId, Integer tcid) {

		String[] stepsLines = dataFormatter.formatCellValue(row.getCell(3)).trim().split("\n");
		String[] expectedLines = dataFormatter.formatCellValue(row.getCell(4)).trim().split("\n");

		if (stepsLines.length != expectedLines.length) {
			return false;
		}

		List<TestCaseStep> steps = new ArrayList<TestCaseStep>();

		for (int i = 0; i < stepsLines.length; i++) {

			TestCaseStep step = new TestCaseStep();
			step.setNumber(i + 1);
			step.setExpectedResults(expectedLines[i]);
			step.setExecutionType(ExecutionType.MANUAL);
			step.setActions(stepsLines[i]);
			steps.add(step);
		}

		TestImportance importance = TestImportance.valueOf(dataFormatter.formatCellValue(row.getCell(5)).toUpperCase());

		TestCase tc = new TestCase(tcid, dataFormatter.formatCellValue(row.getCell(0)), suiteId, projectId, "smiriyala",
				dataFormatter.formatCellValue(row.getCell(1)), steps, dataFormatter.formatCellValue(row.getCell(2)),
				TestCaseStatus.FINAL, importance, ExecutionType.MANUAL, null, null, null, null, false, null, null, null,
				null, null, null, null, null);
		api.updateTestCase(tc);

		return true;

	}

}
