
%% Load data
wdir = "/media/storage/randomisationParkProReakt/";
datafile = "random.xlsx";

data3 = xlsread(fullfile(wdir, datafile), 'E2:I200');
data3 = data3(:,2:4);
[nd, nv]        =   size(data3);

%% Process data

% Normalisation
max_data3       =   max(data3,[],1);
min_data3       =   min(data3,[],1);
data3           =   (data3-repmat(min_data3,nd,1))./(max_data3-min_data3);
% include ones in the fist column
data3           =   [ones(nd,1),data3];
%
npat_first      =   2*nv;

%
irule           =   1; % codifies the rules 1- deterministic; 2- ACA
npat_first      =   2*nv;
for i=1:npat_first
    if (mod(i,2)==0)
        alloc(i)    =   -1;
    else
        alloc(i)    =   1;
    end
end
%

nv              =   nv+1;



fim_groups(1:nv,1:nv,1:2)   =   0;
for i=1:npat_first
    % compute the FIM for ith patient of the first npat_first
    fim_indiv(1:nv,1:nv)    =   transpose(data3(i,1:nv))*data3(i,1:nv);
    if (alloc(i)==-1)
        fim_groups(1:nv,1:nv,1)     =   fim_groups(1:nv,1:nv,1)+ ...
            fim_indiv(1:nv,1:nv);
    else
        fim_groups(1:nv,1:nv,2)     =   fim_groups(1:nv,1:nv,2)+ ...
            fim_indiv(1:nv,1:nv);
    end
end

% count the number of patients already allocated to each treatment
%
% count the number of patients already allocated to each treatment
[iq]                                =   find(alloc(1:npat_first)==-1);
nalloc(1)                           =   size(iq,2);
nalloc(2)                           =   npat_first-nalloc(1);
fim_total(1:nv,1:nv)                =   0;
for i=1:2
    fim_total(1:nv,1:nv)            =   fim_total(1:nv,1:nv)+nalloc(i)/npat_first* ...
        fim_groups(1:nv,1:nv,i);
end

%
bt(1:nv)                            =   0;
for i=1:npat_first
    bt(1:nv)                        =   bt(1:nv)+data3(i,1:nv)*alloc(i);
end
%

% until this point the procedure is only for initializing
% now we start with regular application, i.e. application to each new
% patient arriving
% 
% each time a new patient arrives a new read of the data basis or just of
% the last record must be done.
% Here I ommitted this read as the original data set contains all patients
%
for i=npat_first+1:nd
    inv_fim                         =   inv(fim_total);
    % distance
    d1                              =   data3(i,1:nv)*inv_fim*transpose(bt);
    % rule
    rule                            =   0.5-d1/(1.0+d1^2);
    %
    % now the allocation
    if(irule == 1) % random allocation rule
        ran_val                     =   rand; % randomization
        if ran_val <= 0.5
            alloc(i)                =   1;
        else
            alloc(i)                =   -1;
        end
    else % ACA rule
        ran_val                     =   rand;% randomization
        if ran_val< rule
            alloc(i)                =   1;
        else
            alloc(i)                =   -1;
        end

    end  
    % update variables for a new iteration
    fim_total                   =   double(i-1)/double(i)*fim_total+ ...
        1.0/double(i)*transpose(data3(i,1:nv))*data3(i,1:nv);
    bt(1:nv)                    =   bt(1:nv)+alloc(i)*data3(i,1:nv);  
end
%
