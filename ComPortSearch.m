for i = 1:6
    NoError = 1;
    comStr = ['COM' num2str(i)]
    dmm = serial(comStr,'BaudRate',115200,'Terminator','CR');
    try
        fopen(dmm);
    catch err
        results(i) = err;
        NoError = 0;
    end
    if NoError
        results(i) = 'No Error';
    end
    fclose(dmm);
end