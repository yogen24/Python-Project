from flask import Flask, render_template,redirect, url_for,request,json

import BI_Class
app = Flask(__name__)

@app.route("/",methods=['GET','POST'])
def index():
    return render_template('index.html')

@app.route('/filters/',methods=['POST'])
def my_link():
  array = request.json['get']
  
  if((array[0]=='Region') and (array[1]=='YTD')):
    BI_Class.region_orders()
  if((array[0]=='District') and (array[1]=='YTD')):
    BI_Class.district_orders()
  if((array[0]=='Bank') and (array[1]=='Years')):
    BI_Class.trans_yearly()
  if((array[0]=='Bank') and (array[1]=='Months')):
    BI_Class.trans_monthly()
  if((array[0]=='Bank') and (array[1]=='YTD')):
    BI_Class.bank_orders()
    
  print(array)
  array=[]
  return "x"
   
if __name__=='__main__':
    app.run(debug=True)

