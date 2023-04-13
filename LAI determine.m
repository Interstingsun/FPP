clear
fitdata=[xlsread('C:/Users/Ting/Desktop/watermelon datasets/Run/June.xlsx')]';
VPD=fitdata(2,1:end);
RAD=fitdata(3,1:end);

t=1;
for j=1:18
for i=4:28 
pos=find(fitdata(1,1:end)==j);
         VPDdaily=fitdata(2,pos);
         RADdaily=fitdata(3,pos);
         Ob_Trdaily=fitdata(i,pos);
         [xData, yData, zData] = prepareSurfaceData( VPDdaily, RADdaily, Ob_Trdaily );
    ft = fittype( '(a*(1-exp(-0.86*c))*y+b*c*x)/2.45/1000000*1000*60*0.12^2', 'independent', {'x', 'y'}, 'dependent', 'z' );
    opts = fitoptions( 'Method', 'NonlinearLeastSquares' );
    opts.Display = 'Off';
    opts.Lower = [0 0 0];
    opts.StartPoint = [0.23 22 1.5];
    opts.Upper = [5 100 7];
    [fitresult, gof] = fit( [xData, yData], zData, ft, opts );
    sim_Trdaily=(fitresult.a*(1-exp(-0.86*fitresult.c))*RADdaily+fitresult.b*fitresult.c*VPDdaily)/2.45/1000000*1000*60*0.12^2;
    output=[1:length(Ob_Trdaily);Ob_Trdaily;sim_Trdaily]';
%     set (figure(t),'color','white');
%     plot(output(:,1),output(:,2),output(:,1),output(:,3));  
    Observation(t,:)=[j;i-3;Ob_Trdaily'];
    Simulation(t,:)=[j;i-3;sim_Trdaily'];
    parameter(t,:)=[j;i-3;fitresult.a;fitresult.b;fitresult.c;gof.rsquare]'; 
    parameter_a(j,i-3)=[fitresult.a]; 
    parameter_b(j,i-3)=[fitresult.b]; 
    parameter_c(j,i-3)=[fitresult.c];
t=t+1;
end
end
xlswrite('C:/Users/Ting/Desktop/watermelon datasets/Run/watermelon par_a.xls',parameter_a);
xlswrite('C:/Users/Ting/Desktop/watermelon datasets/Run/watermelon par_b.xls',parameter_b);
xlswrite('C:/Users/Ting/Desktop/watermelon datasets/Run/watermelon par_c.xls',parameter_c);