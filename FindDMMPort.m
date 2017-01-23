%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% FindDMMPort
% This function first calls getAvailableComPort to find out what COM ports
% are available. If any are found it then queries each one until it finds
% the one that has a Fluke 289 DMM on it. If it fails to find either a COM
% port or a Fluke 289 it returns an error.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function dmmPort = FindDMMPort()
    % get the list of available COM ports
    comlist = getAvailableComPort;
    % if the list is not empty proceed
    if ~isempty(comlist)
        listsize = size(comlist);
        % scroll through the available ports
        for i = 1:listsize(1)
            % set up the port to the parameters the Fluke 289 uses
            PortToTry = serial(comlist(i),'BaudRate',115200,...
                               'Terminator','CR');
            PortFound = 1;
            % try to open the port
            try
                fopen(PortToTry);
            catch
                PortFound = 0;
            end
            if PortFound == 1
                try
                    % send the ID command
                    fprintf(PortToTry,'ID');
                catch
                    PortFound = 0;
                end
                if PortFound == 1
                    try
                        % get the ack from the DMM (if it is there)
                        ack = fscanf(PortToTry);
                    catch
                        PortFound = 0;
                    end
                    if isempty(ack)
                        PortFound = 0;
                    else
                        try
                            % get the response from the DMM
                            reading = fscanf(PortToTry);
                        catch
                            PortFound = 0;
                        end
                    end    
                    try
                        % close the port
                        fclose(PortToTry);
                    catch
                        PortFound = 0;
                    end
                end
                if PortFound == 1
                    % check for the correct ID string
                    FlukeIdx = findstr(reading,'FLUKE 289');
                    % if the correct string is found return the COM port address
                    if ~isempty(FlukeIdx)
                        dmmPort = comlist(i);
                        return;
                    else
                        PortFound = 0;
                    end
                end
            end
        end
        % if we complete the list without returning something is wrong
        if i == listsize(1)
            error('Fluke 289 not found')
        end
    else % list is empty return an error
        error('No COM ports available')
    end
end