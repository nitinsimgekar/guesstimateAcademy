package com.guesstimate.utils;

import org.reflections.Reflections;
import org.testng.annotations.Test;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Set;

public class TestCasesCount {

    int total = 0;

    public TestCasesCount() {

        Package[] packages = Package.getPackages();

        for (Package pack : packages) {
            if(pack.getName().contains("testcases")) {
                Reflections reflections = new Reflections(pack.getName());
                Set<Class<? extends GuesstimateBase>> classes = reflections.getSubTypesOf(GuesstimateBase.class);
                for (Class<?> claz : classes) {

                    getTestCaseCount(claz.getName());

                }
            }
        }

        System.out.println("\nTotal Automated: " + total);
    }


    public void getTestCaseCount(String className) {
        try {

            Class<?> cls = Class.forName(className);

            List<Method> methods = new ArrayList<>();
            final List<Method> allMethods = new ArrayList<>(Arrays.asList(cls.getDeclaredMethods()));
            for (final Method method : allMethods) {
                if (method.isAnnotationPresent(Test.class)) {
                    methods.add(method);
                }
            }

            int length = methods.size();

            System.out.println("Test cases count for '"
                    + className.substring(className.lastIndexOf(".") + 1, className.length()) + "': " + length);

            total = total + length;

        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
    }

}
