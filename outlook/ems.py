
class EMS(object):
    def functions(self,weight):
        if 0 <= weight <= 100:
            return {"air":9.2,"ems":21.6}

        elif 101 <= weight <= 200:
            return {"air":12.4,"ems":21.6}

        elif 201 <= weight <=300:
            return {"air":15.6,"ems":21.6}

        elif 301 <= weight <=400:
            return {"air":18.8,"ems":21.6}

        elif 401 <= weight <=500:
            return  {'air':0,"ems":21.6}

        elif 501 <= weight <=600:
            return {"air":25.2,"ems":29.2}

        elif 601 <= weight <=700:
            return {"air":28.4,"ems":29.2}

        elif 701 <= weight <=1000:
            return {'air':0,"ems":29.2}

        elif 1001 <= weight <=1500:
            return {'air':0,"ems":36.8}

        elif 1501 <= weight <=2000:
            return {'air':0,"ems":44.4}

        elif 2001 <= weight <=2500:
            return {'air':0,"ems":52}

        elif 2501 <= weight <=3000:
            return {'air':0,"ems":59.6}

        elif 3001 <= weight <=3500:
            return {'air':0,"ems":67.2}

        elif 3501 <= weight <=4000:
            return {'air':0,"ems":74.8}

        elif 4001 <= weight <=4500:
            return {'air':0,"ems":82.4}

        elif 4501 <= weight <=5000:
            return {'air':0,"ems":90}

        elif 5001 <= weight <=5500:
            return {'air':0,"ems":97.6}

        elif 5501 <= weight <=6000:
            return {'air':0,"ems":105}

        elif 6001 <= weight <=6500:
            return {'air':0,"ems":113}

        elif 6501 <= weight <=7000:
            return {'air':0,"ems":120}

        elif 7001 <= weight <=7500:
            return {'air':0,"ems":128}

        elif 7501 <= weight <=8000:
            return {'air':0,"ems":136}

        elif 8001 <= weight <=8500:
            return {'air':0,"ems":143}

        elif 8501 <= weight <=9000:
            return {'air':0,"ems":151}

        elif 9001 <= weight <=9500:
            return {'air':0,"ems":158}

        elif 9501 <= weight <=10000:
            return {'air':0,"ems":166}

        elif 10001 <= weight <=10500:
            return {'air':0,"ems":174}

        elif 10501 <= weight <=11000:
            return {'air':0,"ems":181}

        elif 11001 <= weight <=11500:
            return {'air':0,"ems":189}

        elif 11501 <= weight <=12000:
            return {'air':0,"ems":196}

        elif 12001 <= weight <=12500:
            return {'air':0,"ems":204}

        elif 12501 <= weight <=13000:
            return {'air':0,"ems":212}

        elif 13001 <= weight <=13500:
            return {'air':0,"ems":219}

        elif 13501 <= weight <=14000:
            return {'air':0,"ems":227}

    def run(self):
        a = self.functions(23.2)
        print(a)

    def main(self):
        self.run()

if __name__=='__main__':
    ems = EMS()
    ems.main()