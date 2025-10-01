import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import accuracy_score

data=pd.read_csv("spam_dataset.csv")

x=data[["number_of_special_characters","number_of_links"]]
y=data[["spam"]]

x_train,x_test,y_train,y_test=train_test_split(x,y,test_size=0.2,random_state=101)

model=DecisionTreeClassifier()
model.fit(x_train,y_train)

y_pred=model.predict(x_test)
# print("accuracy =",accuracy_score(y_test,y_pred))

def predict_op(dt):
    op=model.predict(dt)
    if op[0]==0:
        return 0
    else:
        return 1