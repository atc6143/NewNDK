package com.example.lib;

import org.testng.annotations.Test;

public class AutoTest1 {



    @Test
    public void SanityTest_019() {
        int a = 2, b = 3;
        int sum = 0;
        sum = a + b;
        System.out.println(sum);
    }

    @Test
    public void SanityTest_003() {
        int a = 2, b = 3;
        int sum = 0;
        sum = a + b;
        if(sum>0){
            assert  false:"sum大于0";
        }
    }
    @Test(invocationCount = 0)
    public void SanityTest_010() {
        int a = 2, b = 3;
        int sum = 0;
        sum = a + b;
        System.out.println(sum);
    }
}
