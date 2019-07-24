#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jul 24 14:05:23 2019

@author: rosechrin
"""


def duration(payment_periods, years_to_maturity, ytm, face_value, coupon):
    frequency = payment_periods
    time = years_to_maturity
    ytm = ytm
    value = face_value
    coupon = coupon
    
    pv = ((value*coupon/frequency*(1-(1+ytm/frequency)**(-frequency*time)))/(ytm/frequency)) + \
           value*(1+(ytm/frequency))**(-frequency*time)
           
    final_year = (time*value)/(pv*(1+(ytm/frequency))**(frequency*time))
    
    total = 0
    count = 1

    while count <= frequency*time:
        tn = count/frequency
        count = count + 1
        dur = (value*coupon*tn)/(pv*frequency*(1+(ytm/frequency))**(frequency*tn))
        total = total + dur
       
    print ('Macaulay Duration :' +  str(total+final_year))
    print ('Modified Duration : ' + str((total+final_year)/(1+ytm/frequency)))

