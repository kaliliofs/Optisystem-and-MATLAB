clear all;

NrOfInt         = 255;          % Number of integers per frame
M               = 8;            % Number of bits per symbol
NrFrames        = 10;           % Number of frames to be transmitted

% Generate the matrix with integers to be transmitted
IntSignalin     = randint(NrOfInt,NrFrames,2^M-1);

% Converte integers to binary numbers
for swint = 1:NrFrames
    BinarySig(swint,:)       = reshape( dec2bin(IntSignalin(:,swint)) , NrOfInt * M, 1 );
end

% Setup bit rate and time window for tranmission of each frame
BitRate         = 2.48832e9;    % bits/s
NrOfBits        = NrOfInt*M + 2;
TimeWindow      = NrOfBits / BitRate; 

% Calculate global parameters of OptiSystem to receive the bit sequence
% from Matlab
GlobalNrOfBits  = 1024;
GlobalBitRate   = GlobalNrOfBits / TimeWindow;

% OptiSystem cosimulation -------------------------------------------------

% create a COM server running OptiSystem
optsys = actxserver('optisystem.application');

% This pause is necessary to allow OptiSystem to be opened before
% call open the project
pause(15);

% Open the OptiSystem file defined by the path
optsys.Open('C:\Program Files\Optiwave Software\OptiSystem 10\samples\Matlab cosimulation/OpticalLinkProject.osd');

% Specify and define the parameters that will be varied
ParameterName1  = 'Filename';
InputSignal     = 'OptiSysSequence.dat';

Document        = optsys.GetActiveDocument;
LayoutMngr      = Document.GetLayoutMgr;
CurrentLyt      = LayoutMngr.GetCurrentLayout;
CurrentSweep    = CurrentLyt.GetIteration;
Canvas          = CurrentLyt.GetCurrentCanvas;

% Set the global parameters correctly
CurrentLyt.SetParameterValue('Bit rate', GlobalBitRate);
CurrentLyt.SetParameterValue('Simulation window', 'Set time window');
CurrentLyt.SetParameterValue('Time window', TimeWindow);

% Specify the components that will have the parameters updated
Component1      = Canvas.GetComponentByName('User Defined Bit Sequence Generator');
Component1.SetParameterValue( 'Bit rate', BitRate/1e9 )

Component2      = Canvas.GetComponentByName('Low Pass Bessel Filter');
Component2.SetParameterValue('Cutoff frequency', 0.75 * BitRate/1e9);

Component3      = Canvas.GetComponentByName('Data Recovery');
Component3.SetParameterValue('Reference bit rate', BitRate/1e9);

Visualizer1     = Canvas.GetComponentByName('Binary Sequence Visualizer');

for swint = 1 : NrFrames
    OptiSysSequence(:,1) = str2num(BinarySig(swint,:)');
    
    save OptiSysSequence.dat OptiSysSequence -ascii;
    
    % vary the parameters, run OptiSystem project and get the results
    % Set component parameters                 
    Component1.SetParameterValue( ParameterName1, InputSignal );

    % Calculate                                                                                                                                                                                                                                                                                                 
    Document.CalculateProject( false , true);

    % get the bit sequence recovered from OptiSystem
    GraphBinary      = Visualizer1.GetGraph('Amplitude');
    nSize            = GraphBinary.GetNrOfPoints;

    arrSig           = GraphBinary.GetYData( CurrentSweep );

    SignalOut        = cell2mat( arrSig );
    
    BinaryOut(:,swint) = SignalOut(3:2:2*(NrOfBits-1),1);
    BinaryIn(:,swint)  = OptiSysSequence(:,1);
    BinaryMat       = reshape(BinaryOut(:,swint), M, NrOfInt )';
    % convertes the binary data to integer
    Intmsg          = bin2dec(num2str(BinaryMat));
end

% Calculates the bit error rate and the number of errors detected
[number,ratio] = symerr(BinaryOut,BinaryIn);

disp('The BER calculated is:')
disp(num2str(ratio))
% close OptiSystem
optsys.Quit;