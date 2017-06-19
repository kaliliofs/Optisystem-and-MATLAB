clear all;
close all;

% create a COM server running OptiSystem
optsys = actxserver('optisystem.application');

% This pause is necessary to allow OptiSystem to be opened before
% call open the project
pause(15);

% Open the OptiSystem file defined by the path
optsys.Open(['C:\Users\mverreault\Desktop\OptiSystem 12 Samples\Software interworking\MATLAB co-simulation\MATLAB call OptiSystem\Matlab Call OptiSystem.osd']);

% Specify and define the parameters that will be varied
ParameterName1  = 'Power';
SignalPower     = -20:5:-10; %dBm

ParameterName2 = 'Length';
FiberLength = 5:5:15;   %meters

% Specify the results that will be transfered from OptiSystem
ResultName1     = 'Max. Gain (dB)';
ResultName2     = 'Min. Noise Figure (dB)';
ResultName3     = 'Output : Max. OSNR (dB)';

Document        = optsys.GetActiveDocument;
LayoutMngr      = Document.GetLayoutMgr;
CurrentLyt      = LayoutMngr.GetCurrentLayout;
Canvas          = CurrentLyt.GetCurrentCanvas;

% Specify the components that will have the parameters (results) transfered
Component1      = Canvas.GetComponentByName('CW Laser');
Component2      = Canvas.GetComponentByName('EDFA');
Visualizer1     = Canvas.GetComponentByName('Dual Port WDM Analyzer');

% vary the parameters, run OptiSystem project and get the results
for i = 1:length(SignalPower) 
    for k = 1:length(FiberLength)      
                          
        %Set component parameters                 
        Component1.SetParameterValue( ParameterName1, SignalPower(i) );
        Component2.SetParameterValue( ParameterName2, FiberLength(k) );
           
        %Calculate                                                                                                                                                                                                                                                                                                 
        Document.CalculateProject( true , true);

        %Acces visualizer results
        Result1 = Visualizer1.GetResult( ResultName1 );
        Result2 = Visualizer1.GetResult( ResultName2 );
        Result3 = Visualizer1.GetResult( ResultName3 );
        
        Gain( (i-1)*length(FiberLength) + k )   = Result1.GetValue( 1 );
        NF( (i-1)*length(FiberLength) + k )     = Result2.GetValue( 1 );
        OSNR( (i-1)*length(FiberLength) + k )   = Result3.GetValue( 1 );        
    end
end

%plot graphs
figure
subplot(3,1,1); plot(FiberLength,Gain(1:length(FiberLength)),FiberLength,Gain(length(FiberLength)+1:2*length(FiberLength)),FiberLength,Gain(2*length(FiberLength)+1:3*length(FiberLength)) )
title('Signal Gain')
xlabel('Fiber length [m]')
ylabel('Gain [dB]')
subplot(3,1,2); plot(FiberLength,NF(1:length(FiberLength)),FiberLength,NF(length(FiberLength)+1:2*length(FiberLength)),FiberLength,NF(2*length(FiberLength)+1:3*length(FiberLength)) )
title('Noise Figure')
xlabel('Fiber length [m]')
ylabel('NF [dB]')
subplot(3,1,3); plot(FiberLength,OSNR(1:length(FiberLength)),FiberLength,OSNR(length(FiberLength)+1:2*length(FiberLength)),FiberLength,OSNR(2*length(FiberLength)+1:3*length(FiberLength)) )
title('OSNR')
xlabel('Fiber length [m]')
ylabel('OSNR [dB]')

% close OptiSystem
optsys.Quit;