# -*- coding: utf-8 -*-
"""
Created on Mon Oct 26 13:15:51 2015
Function: Generate SQLite Database Of English Dictionary
@Author: lanbing510
"""

import os
import re
import sqlite3
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

os.chdir(sys.path[0])

db=sqlite3.connect("english_dictionary_sqlite.db")
db.text_factory=str
cu=db.cursor()


# generate sqlite dataset from english_dictionary.txt
cu.execute("create table if not exists english_dictionary (word varchar(50) primary key UNIQUE,explain varchar(200))")
count=0
except_lines=[]
except_exist_lines=[]
except_indexerror_lines=[]
file=open(r"english_dictionary.txt","r")
flog=open(r"log.txt","w+")
for line in file.xreadlines():
    print count
    count+=1
#    if count>1000:
#        break;
    try:
        word=re.findall(r"\s\('(.+)','",line)[0]
        explain=re.findall(r",'(.+)'\);",line)[0]
        command="insert into english_dictionary values('%s', '%s')" % (word,explain)
        cu.execute(command)
        db.commit()
    except sqlite3.IntegrityError,e:
        except_exist_lines.append(count)
        logstr="existed error: %d  %s" % (count, line)
        flog.write(logstr)
    except IndexError,e:
        except_indexerror_lines.append(count)
        logstr="index error: %d  %s" % (count, line)
        flog.write(logstr)
    except Exception,e:
        except_lines.append(count)
        logstr="unknown error: %d  %s" % (count, line)
        flog.write(logstr)
db.close()
file.close()
flog.close()


## refine the items in sqlite dataset
#count=0
#update_count=0
#except_lines=[]
#flog=open(r"log.txt","w+")
#items=cu.execute("select * from english_dictionary")
#items=items.fetchall()
#while count<len(items):
#    print count
#    word_explain=items[count]
#    count+=1;
#    if word_explain[1][0]==" ":
#        update_count+=1;
#        command="update english_dictionary set explain='%s' where word='%s'" % (word_explain[1][1:],word_explain[0])
#        try:
#            cu.execute(command)
#            db.commit()
#        except:
#            except_lines.append(count)
#            flog.write(command+"\n")
#print "have updated %d items of %d items in english dictionary" % (update_count,count)
#db.close()
#flog.close()
