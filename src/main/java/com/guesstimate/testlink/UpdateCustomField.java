package com.guesstimate.testlink;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.Map.Entry;
import java.util.Set;

import org.ini4j.Ini;
import org.ini4j.InvalidFileFormatException;
import org.ini4j.Profile.Section;

import br.eti.kinoshita.testlinkjavaapi.TestLinkAPI;
import br.eti.kinoshita.testlinkjavaapi.model.TestProject;

public class UpdateCustomField {

	public static void main(String[] args) throws InvalidFileFormatException, FileNotFoundException, IOException {

		TestLinkAPI api = UploadTestCases.getInstance(args[0], args[1]);

		System.out.println("Sit back with a irani chai until system updates custom field to testlink.");

		UpdateCustomField updateCustomField = new UpdateCustomField();
		updateCustomField.updateCustomFieldTestCase("./res/testcasexls/" + args[0] + ".ini", api);

	}

	private void updateCustomFieldTestCase(String filePath, TestLinkAPI api)
			throws InvalidFileFormatException, FileNotFoundException, IOException {

		File file = new File(filePath);
		TestProject project = api.getTestProjectByName(file.getName().substring(0, file.getName().lastIndexOf(".")));

		Ini ini = new Ini(new FileReader(filePath));
		Set<Entry<String, Section>> set = ini.entrySet();

		int j = 0;
		int k = 0;

		for (Entry<String, Section> entry : set) {

			Section value = entry.getValue();

			Set<Entry<String, String>> subSet = value.entrySet();

			for (Entry<String, String> entry2 : subSet) {

				String[] split = entry2.getValue().split(",");

				Boolean isCustomFieldAdded = Boolean.valueOf(split[2]);

				if (!isCustomFieldAdded) {

					api.updateTestCaseCustomFieldDesignValue(Integer.valueOf(split[0]), Integer.valueOf(split[1]),
							project.getId(), "AutomationMethodName", entry2.getKey());

					UploadTestCases uploadTestCases = new UploadTestCases();
					uploadTestCases.setTestCaseGenerated(project.getName(), entry.getKey(), entry2.getKey(),
							split[0] + "," + split[1] + ",true");
					j++;

				}

				k++;

			}

		}

		System.out.println("Total " + j + " test cases updated with custom field out of " + k);

	}

}
