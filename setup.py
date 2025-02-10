from iapws import IAPWS97
sat_steam=IAPWS97(P=1,x=1)                #saturated steam with known P
sat_liquid=IAPWS97(T=370, x=0)            #saturated liquid with known T
steam=IAPWS97(h=3462.6, s=7.3251)               #steam with known P and T
print(sat_steam.h, sat_liquid.h, steam.P) #calculated enthalpies
print("hello")