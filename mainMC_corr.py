import numpy as np
import pandas as pd
import scipy.stats as stats
import matplotlib.pyplot as plt
import xlsxwriter
import seaborn as sns

drawNumber = 1000 # Iteration Number

input_FileName = "input.xlsx" # It has to be excel file.
data_SheetName = "Data"
distributions_SheetName = "Distributions"

def ReadFile(FileName):
    
    """
    English:
        
    --> The ReadFile function reads the input.xlsx file which includes input variables. There are one column for average and one column for standart deviation for
    -each input variable.
    
    --> For example, the format of the input.xlsx file should be as follows.
    
            1.column      2.column   3.column         4.column   5.column        . . . n.column
    1.row   Wave Lengths  Average 1  Standart Dev. 1  Average 2  Standart Dev. 2
    2.row      350          0.55      0.45*10^-5       0.7        0.1*10^-10 
    .
    .
    .
    m.row
    
   
    --> After the format is set correctly, it assigns the values of the first column of the input excel file 
    (except for the 1st row because it must be the header row) to the matrix_waveLengths variable.
    -->The size of matrix_waveLengths variable will be ==> Wave Length Number x 1. Their values will be like 350,351....
    -->The size of matrix_values variable will be ==> WaveLength Number  x DataNumber*2
    
    """

    # FileName can be "input_1.xlsx" or "input_2.xlsx"...
    
    # reading first column of input excel file
    df_wavelengths = pd.read_excel(FileName , sheet_name = data_SheetName , usecols = "A" , engine='openpyxl') 
 
    # From the 2nd column (except the 1st column) it reads the input excel file.
    df_values = pd.read_excel(FileName , sheet_name = data_SheetName , index_col=0 , engine='openpyxl')
    
    df_distr = pd.read_excel(FileName , sheet_name = distributions_SheetName , engine = 'openpyxl')
    
    # we need to numpy.array as type. 
    matrix_waveLengths = np.array(df_wavelengths)
    matrix_values = np.array(df_values)
    matrix_distr = np.array(df_distr)
       
    return matrix_waveLengths , matrix_values , matrix_distr

def getTranspose(matrix):
    
    """
    English:
        
    --> matrix must be numpy.array.
    
    --> We are getting transpose of matrix. Because we need to get the transpose of the matrix for some "for" loops.
    
    """
    return matrix.T

def checkIsEqual(mean_array , std_array):
    
    checkMean = np.all(mean_array == mean_array[0])
    checkStd = np.all(std_array == std_array[0])
    
    return checkMean , checkStd

def findDimensions(matrix):
    
    """
    English:
    
    --> This function takes matrix as parameter. Returns the row and column numbers of the matrix.
    
    --> Returning rowNumber and colNumber, our goal is to find Wavelength number and DataNumber with this parameter. 
        
    """
    
    rowNum = len(matrix)
    colNum = len(matrix[0])
     
    return rowNum,colNum

def drawValues(Mean , Stddev , Draws = 1000 , DoF = 1 , Type = "normal"):
    
    if type (Mean) != np.ndarray: # checks if an array of distributions or single distribution is needed
        
        if Type == "normal":
            nxi = stats.norm.rvs (loc=Mean, scale=Stddev, size=Draws)
            return(nxi)
        
        if Type == "T":
            txi=stats.t.rvs (loc=Mean, scale=Stddev, df=DoF, size= Draws)
            return(txi)
       
        if Type == "uniform":
            uxi=np.random.uniform (low=Mean, high=Stddev, size=Draws)
            return(uxi)
        
        if Type == "triangle":
            trxi=stats.triang.rvs(loc=Mean, scale=Stddev, c=DoF, size=Draws)
            return (trxi)
    
    else:
        
        result=np.zeros([len(Mean),Draws])
     
        for i in range (len(Mean)):
            
            if Type == "normal":
                result[i]= stats.norm.rvs (loc=Mean[i], scale=abs(Stddev[i]), size=Draws)
            
            if Type == "T":
                result[i]= stats.t.rvs (loc=Mean[i] , scale=abs(Stddev[i]), df=DoF, size= Draws) 
            
            if Type == "uniform":
                result[i]= np.random.uniform (low=Mean[i], high=Stddev[i], size=Draws)
            
            if Type == "triangle":
                result[i]=stats.triang.rvs(loc=Mean[i], scale=abs(Stddev[i]), c=DoF, size=Draws)
        
        return (result)
    
def sumMC(InputValues, Coverage=0.95, printOutput=False):
    
    #Sorting of the input values
    Ys=sorted(InputValues)
    
    #Calculating the number of draws
    Ylen=len(InputValues)
    
    #Calculating the number of draws covering the given coverage
    q=int(Ylen*Coverage)
    
    #Calculating the draw representing the lower coverage interval boundary
    r= int(0.5*(Ylen-q))
    
    #Calculating the mean of the input values
    ymean=np.mean(InputValues)
    
    #Calculating standard deviation of the input values as absolute standard uncertainty
    yunc=np.std(InputValues)
    
    #Summarizing mean and uncertainty
    values=[ymean,yunc]
    
    #Calculating the values of the draws for olwer and upper boundary of the coverage interval
    ylow=Ys[r]
    yhigh=Ys[r+q]
    
    #Summarizing the coverage intervall
    interval=[ylow,yhigh]
    
    #Summarizing the total output
    output=[values,interval]
    
    #Printing the output values
    if printOutput==True:
        print('mean;'+str(values[0]))
        print('standard deviation:'+str(values[1]))
        print(str(Coverage*100)+'% interval:'+str (interval))
    
    # Returns the output values
    return output

#output2
def calculateMeanStdMC(matrix_Values , matrix_Distr , Data_Number , WaveLength_Number):
    
    """
    English:
    
    --> This function takes matrix_Values , Data_Number and WaveLength_Number as parameters.
    matrix_Values parameter is transpose of matrix of ReadFile function and given as a parameter to the this function.
    Transpose operation is required. Because the rows of our matrix should have mean and standard deviation values, and the column number should be as much as the number of wavelengths.
    For Exemple:
        matrix.Values:
                    1.column(350)       2.column(351)         3.column ...    WaveLength_Number. column
            1.row   average 1       average 1
            2.row   Standart Dev.1  Standart Dev.1
            3.row   average 2       average 2
            .
            .
            .
            Data_Number*2.row
    
    So the size of matrix_Values parameter should be ==> DataNumber*2 x WaveLength Number. It is used accordingly in our loops.
    
    --> The main purpose of this function is to prepare the output2 of the output Excel file.
        
    --> Iterations of the data are taken separately for each wavelength, and the mean and standard deviation values of the iteration matrices are calculated.
        For Example:
            
                 First Data(A1)                     Second Data(A2)   . . .DataNumber.Data(An)
                 
                 A1,350,first iteration             A2,350,first iteration
                 A1,350,second iteration            A2,350,second iteration
                 .
        350:     .
                 .
                 A1,350, drawNumber. iteration      A2,350,drawNumber. iteration
                ------------------------------      ---------------------------
                 A1,350,mean --> mean_matrix[0][0]  A2,350,mean --> mean_matrix[0][1]  
                 A1,350,std  --> std_matrix[0][0]   A2,350,std  --> std_matrix[0][1]
         
         .   
         .
         .
         
    Lastest Wave Length: . . .
    
    --> Calculated values are kept in mean_matrix and std_matrix variables.
    The dimensions of mean_matrix and std_matrix are ==> WaveLength_Number x DataNumber.
    
    -->  It is used in the drawValues and sumMC functions written by PTB in the this function.
    
    
    Changed functionality:
        For the Data A7 to A13 there is only one set of "drawNumber" results generated.
        A7 to A13 do not change over wavelength in the provided input data.
        If one measurement for each quantity A7 to A13 determines the value used for each wavelength, the expected value of the distribution is the same.
        This results in complete correlation over wavelength, meaning the generated random number has to change over each iteration, but has to stay the same for each wavelength.
            
    """
    
    result_list = [] # We use the DrawValues function to hold the returns. and we send to calculateOutputwithFormulMC function.
     
    mean_matrix   = [[0 for j in range(Data_Number)] for i in range(WaveLength_Number)]
    std_matrix    = [[0 for j in range(Data_Number)] for i in range(WaveLength_Number)]
  
    for i in range(0,Data_Number*2,2): # matrix_Values's row number equals Data_Number*2 
        
        data_count = int(((i+1)/2)+1)
        
        check_mean , check_std = checkIsEqual(matrix_Values[i] , matrix_Values[i+1])
        # matrix_Values[i] = average 1 , matrix_Values[i+1] = standart dev. 1.
        if check_mean & check_std == False:
            result = drawValues(matrix_Values[i] , matrix_Values[i+1] , drawNumber , DoF = 1 , Type = matrix_Distr[data_count-1])
            result_list.append(result)
            
            for j in range(0,WaveLength_Number):
            
                output = sumMC(result[j] , Coverage = 0.95 , printOutput= False)
            
                mean_matrix[j][data_count-1] = output[0][0]
                std_matrix[j][data_count-1] = output[0][1]
                
        elif check_mean & check_std == True:
            result = drawValues(matrix_Values[i,0] , matrix_Values[i+1,0] , drawNumber, DoF = 1 , Type = matrix_Distr[data_count-1])
            temp=[]
            for k in range(WaveLength_Number):
                temp.append(result)
            temp_np=np.array(temp)
            result_list.append(temp_np)
            
        # We calculated mean and standart dev. of draws(Iteration Number) for each wave length
            for j in range(0,WaveLength_Number):
                mean_matrix[j][data_count-1] = np.full_like(output[0][0] , matrix_Values[i,0])
                std_matrix[j][data_count-1] = np.full_like(output[0][1] , matrix_Values[i+1,0])
    
    
    return result_list , mean_matrix , std_matrix

#output1
def calculateOutputwithFormulaMC(result_list , Data_Number , WaveLength_Number):
    
    """
   English:
   --> The main purpose of this function is to prepare the "output1" result of the "output" excel file for Monte Carlo iteration results.
   
   --> This function takes "result_list" , "Data_Number" and "WaveLength_Number" as parameters.
   
   --> "result_list": It corresponds to the "result_list" in the return of the "calculateMeanStdMC" function. 
       The length of the "result_list" is expected to be as much as the "DataNumber" number.
  
   --> The values produced as a result of iteration of each data are used in the formula function of the ones belonging to the same index and the result is obtained
   and this is applied for each wavelength.
  
   For Example:       
        drawNumber(iterasyon number) = 4
        Data_Number = 2
        WaveLength_Number = 3
        formula equation : I = A1+ A2 
        
         A1          A2
        A1,350,1   A2,350,1 -->    I350,1 = A1,350,1 + A2,350,1
        A1,350,2   A2,350,2 -->    I350,2 = A1,350,2 + A2,350,2
  350:  A1,350,3   A2,350,3 -->    I350,3 = A1,350,3 + A2,350,3
        A1,350,4   A2,350,4 -->    I350,4 = A1,350,4 + A2,350,4
                               ------------------------------
                                   I350,mean , I350,std
                                   
        A1          A2
        A1,351,1   A2,351,1 --> I351,1 = A1,351,1 + A2,351,1
        A1,351,2   A2,351,2 --> I351,2 = A1,351,2 + A2,351,2
  351:  A1,351,3   A2,351,3 --> I351,3 = A1,351,3 + A2,351,3
        A1,351,4   A2,351,4 --> I351,4 = A1,351,4 + A2,351,4
                               ------------------------------
                                   I351,mean , I351,std
    
                                      
        A1          A2
        A1,352,1   A2,352,1 --> I352,1 = A1,352,1 + A2,352,1
        A1,352,2   A2,352,2 --> I352,2 = A1,352,2 + A2,352,2
  352:  A1,352,3   A2,352,3 --> I352,3 = A1,352,3 + A2,352,3
        A1,352,4   A2,352,4 --> I352,4 = A1,352,4 + A2,352,4
                               ------------------------------
                                   I352,mean , I352,std
  
       
     --> "output_matrix" keeps the mean and standard deviation values that we have calculated for each wavelength.
         The size of "output_matrix" is ==> WaveLength_Number x 2.  
  
   --> "draw_matrix" is the creation of a matrix from the common index elements of random values as much as the number of iterations produced for each data.
   That is, it is rendered like draw_matrix[0]= [A1,350.1 , A2,350.1].
   size of "draw_matrix": drawNumber x Data_Number
  
    
   Changed functionality:
       
   --> "mc_values" is added as a matrix for keeping the monte carlo results after using the drawn values in the formula.
   The "draw_matrix" values are used in the "formula" so for each wavelength there are "iteration_number" results of the formula.
   This are required for calculation correlations.
   
   """
    
    output_matrix = [[0 for j in range(2)] for i in range(WaveLength_Number)]
    draw_matrix   = [[0 for j in range(Data_Number)] for i in range(drawNumber)]
    mc_Values = np.zeros((drawNumber, WaveLength_Number))
          
    # We calculated formul result's mean and standart dev. of a wave length until draws number(iteration number.)
    for k in range(0,WaveLength_Number):  
        
        for i in range(0,drawNumber):
            
            for j in range(0,Data_Number): 
            
                temp_result = result_list[j] 
                # We create the inside of the draw_matrix.
                draw_matrix[i][j] = temp_result[k][i] 
        
        # The created draw_matrix is sent to the formula and the formula is written in its place with the necessary indexes, the result is calculated and thrown into the output_matrix.                       
        output_matrix[k][0] , output_matrix[k][1] , mc_Values[:,k] = formula(draw_matrix) 
        
    return output_matrix , mc_Values

def formula(draw_matrix):
    
    """
    English:
     --> The main purpose of this function is to calculate and return the mean and standard deviation values by formulating them with 
     the values of the draw_matrix variable created separately for each wavelength.
    --> This function takes the draw_matrix parameter. The formula function is called by the calculateOutputwithFormulaMC function.
    draw_matrix corresponds to the draw_matrix matrix in this function.
    The size of draw_matrix should be ==> drawNumber x DataNumber.
        
    --> Draw_matrix is created for each wavelength. It is calculated by putting the correct places in the formula with draw_matrix.
     For Example:
        draw_matrix[0] = [A1,350,1 , A2,350,1]
        draw_matrix[1] = [A1,350,2 , A2,350,2]
        .
        .
        .
        draw_matrix[drawNumber - 1 ] = [A1,350,drawNumber , A2,350,DrawNumber]
        
        
    
    Changed functionality:
        The complete output_list is returned as a result.
        The list is requiered for correlation analysis.
        By calculating mean and std immediately all information about correlations included in the random numbers is lost.    
   
    """
    # draw_matrix --> drawNumber x Data Number size

    formula = 0.0
    
    output_list = []
    
    for i in range(0,len(draw_matrix)):
        
        # if your data number is not equal to 13, it will calculate 0 for the output1 result. You should define a formula with an if condition for your data set.
        if len(draw_matrix[0]) == 13: # data number --> 13 
        
            formula = (draw_matrix[i][0]+draw_matrix[i][1])*(draw_matrix[i][2]/draw_matrix[i][3])*(draw_matrix[i][4]/draw_matrix[i][5])*(1+draw_matrix[i][6]+draw_matrix[i][7]+draw_matrix[i][8]+draw_matrix[i][9]+draw_matrix[i][10]+draw_matrix[i][11]+draw_matrix[i][12]) 
            
        output_list.append(formula)  
        
    # calculated mean and standart dev. of a wave length . 
    # to access correlations the monte carlo results, without calculation of mean and standarddev has to be preserved
    return np.mean(output_list) , np.std(output_list) , output_list

def writeExcel(output_matrix , mean_matrix , std_matrix , Data_Number , WaveLength_Number , matrix_WaveLengths):

    """
    
    English:
        
    --> The function takes output_matrix , mean_matrix , std_matrix , Data_Number , WaveLength_Number and matrix_WaveLengths as parameters.
    
    output_matrix:calculateOutputwithFormulaMC corresponds to the returned matrix of the function and expects it.
    
    mean_matrix , std_matrix: corresponds to the returned mean_matrix and std_matrix matrices of the calculateMeanStdMC function and expects them.
    
    WaveLength_Number: It corresponds to 1 minus the number of lines of the Excel file. Or it is equal to the returned Column Number of the findDimensions function and is expected.
    
    matrix_WaveLengths: Corresponds to the matrix_WaveLengths returned from the ReadFile function.
    
    --> The results produced as a result of Monte Carlo are written to Excel output_drawNumber.xlsx file.
    
    """
    
    workbook = xlsxwriter.Workbook("output_"+str(drawNumber)+".xlsx") 
    worksheet_formul_output = workbook.add_worksheet("output1")
    worksheet_mean_std_output = workbook.add_worksheet("output2")
   
    
    worksheet_mean_std_output.write(0,0,"Wave Lengths") 
    worksheet_formul_output.write(0,0,"Iteration Number")
    worksheet_formul_output.write(0,1,""+str(drawNumber))
    worksheet_formul_output.write(1,0,"Wave Lengths")
    worksheet_formul_output.write(1,1,"Mean")
    worksheet_formul_output.write(1,2,"Standart Dev")
    
    for i in range(1,WaveLength_Number+1):
        worksheet_mean_std_output.write(i,0,matrix_WaveLengths[i-1])
        worksheet_formul_output.write(i+1,0,matrix_WaveLengths[i-1])
     
    for i in range(1,WaveLength_Number+1):
        
        for j in range(0,Data_Number):
            
            # we are controling;is it first row or not ? true : false
            if i - 1 == 0:
    
                # j == 0 --> column: 1 , 2 j==1 --> column: 3 ,4  j==2 --> column: 5 , 6 
                worksheet_mean_std_output.write(i-1,2*j+1,"mean"+str(j+1))
                worksheet_mean_std_output.write(i-1,2*j+2,"std"+str(j+1))
                
            worksheet_mean_std_output.write(i,2*j+1,mean_matrix[i-1][j])
            worksheet_mean_std_output.write(i,2*j+2,std_matrix[i-1][j])
            
        worksheet_formul_output.write(i+1 , 1 , output_matrix[i-1][0]) # row --> 0 2 4
        worksheet_formul_output.write(i+1 , 2 , output_matrix[i-1][1]) # row --> 1 3 5
    
    print("\nWriting is finished")
    
    workbook.close()
    
def correlation (Distributions):
     
    matrix = np.corrcoef(Distributions)
    return(matrix)
    
def corrPlot(Corr_Matrix , data1_index , data2_index):
    
    """
    English:
        
    --> This function takes Corr_Matrix , data1_index and data2_index as parameters.
    
    Corr_Matrix: It corresponds to the Correlation Coefficient Matrix. 
    If you want to look at the wavelength basis, the size of the Correlation Coefficient Matrix will be ==> WaveLength Number x WaveLength Number.
        
    --> data1_index and data2_index is important if you want to look at the correlation between 2 data. It loses its importance for more than 2 data.
    If you want to see correlation between x1 and x2 data, data1_index = 0 , data2_index = 1 .
    
    """

    if len(Corr_Matrix) == 2:
        
        # To see between data.
        sns.heatmap(Corr_Matrix , vmin = -1 , vmax = 1 , annot = True , xticklabels= [data1_index+1,data2_index+1] , yticklabels=[data2_index+1,data1_index+1])
        
    
    elif len(Corr_Matrix) > 2:
        
        # To see on the basis of wavelength.
        fig = plt.figure()#dpi=1100)
        subplot=fig.add_subplot(111)
        cax=subplot.imshow(Corr_Matrix , vmin=-1 , vmax=1 , cmap="jet" , interpolation="nearest" , origin = "lower")
        fig.colorbar (cax, ticks=[-1,-0.75,-0.5,-0.25,0,0.25,0.5,0.75,1])
        plt.show()
        
    else:
        print("Error!!! You gave the wrong Correlation Coefficient Matrix!!!")

def spectralcorrelation(mc_matrix):
    
    """
    
    Parameters
    ----------
    mc_matrix : Monte carlo values required for calculation of the correlation of the result data.

    Returns
    -------
    None.
    
    """
    # mc_matrix's shape = drawNumber x WaveLength_Number 
    
    corrMatrix=correlation(getTranspose(mc_matrix))   # calculates the correlation of the result MC values with each other
    corrPlot(corrMatrix,"","")
    
    return corrMatrix
    
def resultplot(matrix_WaveLengths , output_matrix):
    
    """
    --> Function to plot the spectral output values and the relative standard deviation as a function of wavelength
    

    Parameters
    ----------
    matrix_WaveLengths : wavelength vector provided for x-axis
    output_matrix : matrix with mean values and std for plotting the data

    Returns
    -------
    None.

    """
    
    outputvector=np.array(output_matrix)
    fig, ax = plt.subplots()
    ax.plot(matrix_WaveLengths, outputvector[:, 0],
        linestyle="None", marker="o", color="blue")
    ax.set_xlabel("wavelength / nm")
    ax.set_ylabel("E")
    ax2 = ax.twinx()
    ax2.plot(matrix_WaveLengths, outputvector[:, 1]/outputvector[:, 0],
         linestyle="None", marker="o", color="red")
    ax2.set_ylabel("u_rel(E)")
    #ax2.set_yscale("log")
    plt.show()    

def createDistributionsforCopulas(mc_matrix):
    
    """
    mc_matrix.size --> (drawNumber , WaveLength_Number)
    We need to "getTranspose"
    
    """
    
    Distributions = []
    transpose_mcMatrix = getTranspose(mc_matrix)
    
    # we are creating Distributions for each Wave Length to calculate copulas. e.g: [[mean , std ,'n'] , [mean , std , 'u] . . .]
    # what is the 'n' and 'u' ? if they are a Distributions shortcuts, you need to a string variable with if conditional.
    for i in range(0,len(transpose_mcMatrix)):
        
        temp_list = []
        temp_list = [np.mean(transpose_mcMatrix[i])  , np.std(transpose_mcMatrix[i]) , 'n']
        Distributions.append(temp_list)
        
    
    return Distributions
    

def drawMultiVariate(Distributions , Correlationmatrix , Draws=1000):
         
        """
        --> Draw values from a multivariate standard distribution according to the given correlation matrix.
                
        --> Returns an array with the dimensions (Number of Distributions,Number of Draws).
                
        Example: drawMultiVariate (List[[Mean, total standard uncertainty,type],...],correlation matrix)
                    
        --> Within the distribution list for type "n" represents standard distribution and "u" represents uniform distribution.
                    
        --> As Distributions a list is needed.
        Example for a standard and uniform distribution: Distribution=[[1,0,1,"n"][5,1,"u"]]
                        
        As Correlationmatrix a positive semidefinite Matrixarray as needed:
        Example for two quantities with correlation rho: numpy.array ([1.0,rho],[rho,1.0])
                   
        """
        
        dimension= len(Correlationmatrix)
        copula = stats.multivariate_normal(np.zeros(dimension),Correlationmatrix)   
        
        z=copula.rvs(Draws)
        #x = [[0 for j in range(Draws)] for i in range(dimension)]
        x=np.zeros(shape=(dimension,Draws))
        
        for i in range (dimension):
            
            xi= stats.norm.cdf(z[:,i])
            
            if Distributions [i][2]=="n":
            
                xidist= stats.uniform.ppf(xi,loc= Distributions[i][0],scale= Distributions[i][1])
            
            if Distributions [i][2]=="u":
                
                xidist= stats.uniform.ppf(xi,loc= Distributions[i][0],scale= Distributions[i][1])
        
            x[i]=xidist
        
        return(x)


def scatterPlotCopulas(x):
    
    # we are looking two between wave length.
    plt.scatter(x[100] , x[200] , drawNumber , color= "blue" , alpha = 0.7)
    plt.title('copula')
    plt.xlabel('Wave Length - value')
    plt.ylabel('Wave Length - value')
    plt.show()
    
    
def mainMC(drawNum,FileName):
    
    """
    English:
        
    --> The function takes 2 parameters as drawNum and FileName as parameters. "DrawNum" corresponds to the number of iterations, 
    while "FileName" corresponds to the name of the "input" excel file.
    These two parameters must be defined as global variables at the top of the Python code and given as parameters when calling this function.
    
    --> "mainMC" function is the function that calls the functions that read  "input" excel file, apply the MonteCarlo method and write results to "output" Excel file.
    
    
    Changed functionality:
        Added "mc_matrix" to be included. 
        Added "version" to enable easy comparison between the old version of generating values and the changed version
        
    """
    
    print("\nMonte Carlo is started\n")
    
    matrix_WaveLengths , matrix_Values , matrix_Distr = ReadFile(FileName)
    matrix_Values = getTranspose(matrix_Values)
    matrix_Distr  = getTranspose(matrix_Distr)
    
    row_Number , Column_Number = findDimensions(matrix_Values)
    
    # Normally, the WaveLength_Number variable will be equal to the Row Number. However, since we transpose it, it equals the number of columns. The reason for doing this is to use it more comfortably in our loops.
    WaveLength_Number = Column_Number
 
    # This condition is only for checking whether the format of the input excel file prepared for Monte Carlo calculation is correct.
    if row_Number % 2 == 0 :
        Data_Number = int(row_Number/2)
    
    else :
        print("row Number has to be even !!!")
   
    
    result_list , mean_matrix , std_matrix = calculateMeanStdMC(matrix_Values , matrix_Distr , Data_Number, WaveLength_Number)  
    output_matrix, mc_matrix = calculateOutputwithFormulaMC(result_list , Data_Number, WaveLength_Number)
    
    print("\nMonte Carlo is finished\n")
    
    writeExcel(output_matrix , mean_matrix, std_matrix , Data_Number , WaveLength_Number, matrix_WaveLengths) 
    
    return output_matrix , mean_matrix , std_matrix, mc_matrix, matrix_WaveLengths

# ------- running ------- 

output_matrix , mean_matrix , std_matrix , mc_matrix , matrix_WaveLengths  = mainMC(drawNumber,input_FileName)
corrMatrix = spectralcorrelation(mc_matrix)
resultplot(matrix_WaveLengths , output_matrix)

Distributions = createDistributionsforCopulas(mc_matrix)
# x's size is same get transpose mc_matrix's size. We are expecting it. x size --> (wave length number ,  drawNumber)
x = drawMultiVariate(Distributions , corrMatrix , drawNumber)
scatterPlotCopulas(x)

#------------------------







