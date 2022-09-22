import os,re,glob
from DataProcess import dataprocess
from StyleProcess import styleprocess

path = os.getcwd()
os.chdir(path)
robot_prefix = ['IRB','CRB','IRBT']

def get_robotlist(robot_prefix_list):

    path=os.getcwd()
    rl=[]

    for robot_prefix in robot_prefix_list:
        robotregex=re.compile(r'{0} .*?.xlsm'.format(robot_prefix))
        
        for file in glob.glob(path + '/*.xlsm'):

            mo1=robotregex.search(file)
            if mo1 and str(mo1.group()).split('.')[0] not in rl:
                rl.append(str(mo1.group()).split('.')[0])
    
    return rl

if __name__ == '__main__':

    robottypes = get_robotlist(robot_prefix)

    for robottype in robottypes:
        try:
            print('{0} DataProcess is running....'.format(robottype))
            dataprocess(robottype,path)
            print('{0} StyleProcess is running....'.format(robottype))
            styleprocess(robottype,path)
        except Exception as ex:
            print(ex)
            print('{0} file generates failed!'.format(robottype))
            os.system('pause')
            print('---------------------------------\n')
        finally:
            continue
    os.system('pause')