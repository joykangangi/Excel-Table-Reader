%reading the table from excel
mcsResults = readtable('MCS_3.1.xlsx'); 
unitsArray = ["ODE","OperationsResearchI","ScientificComputing","ComputerGraphics","Accounts_Finance","NumericalAnalysis","OperatingSystemsII","RealAnalysis"]

%class mean in each unit
unitsMean = varfun(@mean,mcsResults,"InputVariables",unitsArray)

fprintf('The mean in ODE:%.0f\n',unitsMean.mean_ODE);
fprintf('The mean in Operations Research I:%.0f\n',unitsMean.mean_OperationsResearchI);
fprintf('The mean in Scientific Computing:%.0f\n',unitsMean.mean_ScientificComputing);
fprintf('The mean in Computer Graphics:%.0f\n',unitsMean.mean_ComputerGraphics);
fprintf('The mean in Accounts&Finance:%.0f\n',unitsMean.mean_Accounts_Finance);
fprintf('The mean in Numerical Analysis:%.0f\n',unitsMean.mean_NumericalAnalysis);
fprintf('The mean in Operating Systems II:%.0f\n',unitsMean.mean_OperatingSystemsII);
fprintf('The mean in Real Analysis:%.0f\n',unitsMean.mean_RealAnalysis);

%changing table to matrix
unitTitle = vartype('numeric');
unitInNumeric = mcsResults(:,unitTitle);
resultMatrix = unitInNumeric{:,:};


%average of each student/ mean of each row in matrix.
%creates an AVERAGE column for the mean in the table
mcsResults.AVERAGE = mean(resultMatrix,2)

%grade the scores in each unit, the overall grade and the respective remark
mcsResults.ODE = discretize(mcsResults.ODE,[0,40,50,60,70,100],'categorical',{'E','D','C','B','A'});
mcsResults.OperationsResearchI = discretize(mcsResults.OperationsResearchI,[0,40,50,60,70,100],'categorical',{'E','D','C','B','A'});
mcsResults.ScientificComputing = discretize(mcsResults.ScientificComputing,[0,40,50,60,70,100],'categorical',{'E','D','C','B','A'});
mcsResults.ComputerGraphics = discretize(mcsResults.ComputerGraphics,[0,40,50,60,70,100],'categorical',{'E','D','C','B','A'});
mcsResults.Accounts_Finance = discretize(mcsResults.Accounts_Finance,[0,40,50,60,70,100],'categorical',{'E','D','C','B','A'});
mcsResults.NumericalAnalysis = discretize(mcsResults.NumericalAnalysis,[0,40,50,60,70,100],'categorical',{'E','D','C','B','A'});
mcsResults.OperatingSystemsII = discretize(mcsResults.OperatingSystemsII,[0,40,50,60,70,100],'categorical',{'E','D','C','B','A'});
mcsResults.RealAnalysis = discretize(mcsResults.RealAnalysis,[0,40,50,60,70,100],'categorical',{'E','D','C','B','A'});
mcsResults.GRADE = discretize(mcsResults.AVERAGE,[0,40,50,60,70,100],'categorical',{'E','D','C','B','A'});
mcsResults.REMARK = discretize(mcsResults.AVERAGE,[0,40,50,60,70,100],'categorical',{'Fail','Pass','Satisfactory','Good','Excellent'});

%printing Transcripts
for s=1:164
    fprintf("<strong>NAME</strong>:\t%s",char(mcsResults.NAMES(s)))
    fprintf("<strong>REGISTRATION NUMBER</strong>:%s",char(mcsResults.REG_NO(s)))
    fprintf("<strong>ODE</strong>:\t\t\t%s",char(mcsResults.ODE(s)))
    fprintf("<strong>REAL ANALYSIS</strong>:\t\t%s",char(mcsResults.RealAnalysis(s)))
    fprintf("<strong>OPERATIONS RESEARCH</strong>:\t%s",char(mcsResults.OperationsResearchI(s)))
    fprintf("<strong>SCIENTIFIC COMPUTING</strong>:\t%s",char(mcsResults.ScientificComputing(s)))
    fprintf("<strong>COMPUTER GRAPHICS</strong>:\t%s",char(mcsResults.ComputerGraphics(s)))
    fprintf("<strong>ACCOUNTS AND FINANCE</strong>:\t%s",char(mcsResults.Accounts_Finance(s)))
    fprintf("<strong>NUMERICAL ANALYSIS</strong>:\t%s",char(mcsResults.NumericalAnalysis(s)))
    fprintf("<strong>OPERATING SYSTEMS II</strong>:\t%s",char(mcsResults.OperatingSystemsII(s)))
    fprintf("<strong>AVERAGE</strong>:\t\t%.1f",char(mcsResults.AVERAGE(s)))
    fprintf("<strong>OVERALL GRADE</strong>:\t\t%s",char(mcsResults.GRADE(s)))
    fprintf("<strong>REMARK</strong>:\t%s",char(mcsResults.REMARK(s)))
    fprintf(' ','\n')
    fprintf(' ','\n')
end
s=s+1;
