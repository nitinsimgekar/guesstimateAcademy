package com.guesstimate.utils;

import com.github.javaparser.ast.CompilationUnit;
import com.github.javaparser.ast.PackageDeclaration;
import com.github.javaparser.ast.body.ClassOrInterfaceDeclaration;
import com.github.javaparser.ast.expr.Name;
import com.guesstimate.page.BasePageMobile;
import com.guesstimate.page.BasePageWeb;
import org.ini4j.Ini;
import org.ini4j.Profile;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.FindBy;
import org.reflections.Reflections;

import javax.tools.JavaCompiler;
import javax.tools.StandardJavaFileManager;
import javax.tools.ToolProvider;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;

public class GenerateActionsMethods {

    int total = 0;

    public GenerateActionsMethods() {

        Package[] packages = Package.getPackages();

        List<String> failedFiles = new ArrayList<>();

        for (Package pack : packages) {
            Reflections reflections = new Reflections(pack.getName());

            Set<Class<? extends BasePageWeb>> webPageClasses = reflections.getSubTypesOf(BasePageWeb.class);
            for (Class<?> claz : webPageClasses) {

                getVariablesW(claz.getName(), pack.getName(), failedFiles);

            }

            Set<Class<? extends BasePageMobile>> mobilePageClasses = reflections.getSubTypesOf(BasePageMobile.class);
            for (Class<?> claz : mobilePageClasses) {

                getVariablesM(claz.getName(), pack.getName(), failedFiles);

            }

        }

        if (!failedFiles.isEmpty()) {
            System.out.println("\nThe following files failed to generate methods");
            for (String file : failedFiles) {
                System.out.println(file);
            }
        }

        System.out.println("\nTotal action methods generated: " + total);
    }

    private void getVariablesW(String className, String packageName, List<String> failedFiles) {

        try {

            Class<?> cls = Class.forName(className);

            Set<String> keys = getVariableActionGenerated(className);

            Services services = new Services();
            String filePath = "./src/main/java/" + className.replace(".", "/") + ".java";
            String output = services.readFileWithSpaces(filePath).trim();

            if (!output.isEmpty()) {

                output = output.substring(0, output.lastIndexOf("}"));

                if (!output.contains("import org.openqa.selenium.support.ui.ExpectedConditions;")) {
                    String str1 = output.substring(0, output.lastIndexOf("import"));
                    String str2 = output.substring(output.lastIndexOf("import"), output.length());

                    output = str1 + "import org.openqa.selenium.support.ui.ExpectedConditions;\n" + str2;

                }




                StringBuilder sb = new StringBuilder(output);
                sb.append("\n\n");

                ArrayList<Field> allVariables = new ArrayList<>(Arrays.asList(cls.getDeclaredFields()));
                for (Field variable : allVariables) {
                    if (variable.isAnnotationPresent(FindBy.class) && !keys.contains(variable.getName())) {

                        if (variable.getName().endsWith("btn") || variable.getName().startsWith("btn") || variable.getName().endsWith("Btn") || variable.getName().endsWith("Btn")) {

                            sb.append("public " + className.substring(className.lastIndexOf(".") + 1, className.length()) + " click" + variable.getName().substring(0, 1).toUpperCase() + variable.getName().substring(1, variable.getName().length()) + "() {\n" +
                                    "        wait.until(ExpectedConditions.elementToBeClickable(" + variable.getName() + "));\n" +
                                    "        " + variable.getName() + ".click();\n" +
                                    "        return this;\n" +
                                    "    }\n\n");


                        } else if (variable.getName().endsWith("lbl") || variable.getName().startsWith("lbl") || variable.getName().endsWith("Lbl") || variable.getName().endsWith("Lbl")) {

                            sb.append("public String get" + variable.getName().substring(0, 1).toUpperCase() + variable.getName().substring(1, variable.getName().length()).replace("lbl", "").replace("Lbl", "") + "() {\n" +
                                    "        wait.until(ExpectedConditions.visibilityOf(" + variable.getName() + "));\n" +
                                    "        return " + variable.getName() + ".getText();\n" +
                                    "    }\n\n");


                        } else if (variable.getName().endsWith("edit") || variable.getName().startsWith("edit") || variable.getName().endsWith("Edit") || variable.getName().endsWith("Edit")) {

                            String name = variable.getName().replace("edit", "").replace("Edit", "");
                            name = name.substring(0, 1).toUpperCase() + name.substring(1, name.length());

                            sb.append("public " + className.substring(className.lastIndexOf(".") + 1, className.length()) + " enter" + name + "(String " + variable.getName().replace("edit", "").replace("Edit", "").toLowerCase() + "AsString) {\n" +
                                    "        wait.until(ExpectedConditions.visibilityOf(" + variable.getName() + "));\n" +
                                    "        " + variable.getName() + ".sendKeys(" + variable.getName().replace("edit", "").replace("Edit", "").toLowerCase() + "AsString);\n" +
                                    "        return this;\n" +
                                    "    }\n\n");


                        }


                    }
                }

                sb.append("}");

                Formatter formatter = new Formatter();

                FileWriter writer = new FileWriter(filePath);
                writer.write(formatter.format(sb.toString()).toString());
                writer.close();

            } else
                failedFiles.add(className);

        } catch (ClassNotFoundException | IOException e) {
            e.printStackTrace();
        }


    }

    private void getVariablesM(String name, String className, List<String> failedFiles) {
    }

    private Set<String> getVariableActionGenerated(String className) {

        Set<String> keys = null;
        try {
            String iniFileName = "./res/generateActions.ini";
            File file = new File(iniFileName);

            if (!file.exists()) {
                file.createNewFile();
            }

            Ini ini = new Ini(new FileReader(file));
            Profile.Section sections = ini.get(className);

            if (sections == null)
                keys = new HashSet<>();
            else
                keys = sections.keySet();

        } catch (IOException e) {
            e.printStackTrace();
        }

        return keys;
    }

    public void setTestCaseGenerated(String className, String varName) {

        try {
            String iniFileName = "./res/generateActions.ini";
            Ini ini = new Ini(new FileReader(new File(iniFileName)));

            ini.put(className, varName, true);
            ini.store(new File(iniFileName));

        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
