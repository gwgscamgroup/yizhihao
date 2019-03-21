% �����о��ܱ��е����ݲ���
%% ����
endT='2019/3/15';
period='W';
useWsd=1;   % �Ƿ���wsd����wssȡ����


fileName='D:\Job2\�о������ͼ\��ͼ.xlsx';
fileSec='D:\Job2\�о������ͼ\�¸���.xlsx';
codeIndex={'000001.SH';'399006.SZ';'000300.SH';'000016.SH';'000905.SH'};
%% ����
w=windmatlab;
nBack=3;
begT=w.tdaysoffset(-nBack,endT,['Period=',period]);
dateP=w.tdays(begT,endT,['Period=',period]);
while length(dateP)<4
    nBack=nBack+1;
    begT=w.tdaysoffset(-nBack,endT,['Period=',period]);
    dateP=w.tdays(begT,endT,['Period=',period]);
end
dateP(1)=[];
dateDaily=w.tdays(begT,endT);
[~,indP]=ismember(dateP,dateDaily);
sec2=datestr([dateDaily(indP(1)+1);dateP(2)],'yyyymmdd');   % ������
sec1=datestr([dateDaily(indP(2)+1);dateP(3)],'yyyymmdd');   % ����

% ָ������
if useWsd    
    price0=w.wsd(codeIndex,'close',dateP{2},dateP{2});
    price1=w.wsd(codeIndex,'close',dateP{3},dateP{3});    
    high=w.wsd(codeIndex,'high',sec1(1,:),dateP{3});
    low=w.wsd(codeIndex,'low',sec1(1,:),dateP{3});    
    amt=w.wsd(codeIndex,'amt',sec1(1,:),dateP{3});
    
    nameIndex=w.wsd(codeIndex,'sec_name');
    chIdx=(price1./price0-1)*100;
    swIdx=(max(high)'-min(low)')./price0*100;
    amIdx=sum(amt)';
    peIdx=w.wsd(codeIndex,'pe_ttm',dateP{3},dateP{3});
else
    nameIndex=w.wss(codeIndex,'sec_name'); %#ok<UNRCH>
    chIdx=w.wss(codeIndex,'pct_chg_per',['startDate=',sec1(1,:)],['endDate=',sec1(2,:)]);
    swIdx=w.wss(codeIndex,'swing_per',['startDate=',sec1(1,:)],['endDate=',sec1(2,:)]);
    amIdx=w.wss(codeIndex,'amt_per','unit=1',['startDate=',sec1(1,:)],['endDate=',sec1(2,:)]);
    peIdx=w.wss(codeIndex,'pe_ttm',['tradeDate=',sec1(2,:)]);
end
stat=cat(1,cat(2,{nan},nameIndex'),cat(2,{'�ǵ���';'���';'�ɽ���';'��ӯ��'},num2cell([chIdx,swIdx,amIdx,peIdx]')));
disp(stat);

% ��ҵ��
nameIdy=w.wset('sectorconstituent',['date=',endT,';sector=�������һ����ҵָ��;field=sec_name']);
nameIdy=strrep(nameIdy,'(����)','');
numIdy=length(nameIdy);
peIdy=cell(numIdy,3);
chIdy=cell(numIdy,2);
for n=1:numIdy
    peIdy(n,1)=w.wsee(strcat('SW',nameIdy{n}),'sec_pe_ttm_overall_chn',['tradeDate=',sec1(2,:)],'excludeRule=2','DynamicTime=1');
    peIdy(n,2)=w.wsee(strcat('SW',nameIdy{n}),'sec_pe_ttm_overall_chn',['tradeDate=',sec2(2,:)],'excludeRule=2','DynamicTime=1');
    chIdy(n,1)=w.wsee(strcat('SW',nameIdy{n}),'sec_pq_pct_chg_ffmc_wavg_chn',['startDate=',sec1(1,:)],['endDate=',sec1(2,:)],'DynamicTime=1');
    chIdy(n,2)=w.wsee(strcat('SW',nameIdy{n}),'sec_pq_pct_chg_ffmc_wavg_chn',['startDate=',sec2(1,:)],['endDate=',sec2(2,:)],'DynamicTime=1');
    disp(['��ҵ��',num2str(n),'-',num2str(numIdy)]);
end
peIdy=cell2mat(peIdy);
chIdy=cell2mat(chIdy);
peIdy(:,3)=peIdy(:,1)./peIdy(:,2)-1;

% ������
[~,~,nameSec]=xlsread(fileSec);
numSec=length(nameSec);
chSec=cell(numSec,1);
for n=1:numSec
    chSec(n)=w.wsee(nameSec{n},'sec_pq_pct_chg_ffmc_wavg_chn',['startDate=',sec1(1,:)],['endDate=',sec1(2,:)],'DynamicTime=1');
    disp(['���',num2str(n),'-',num2str(numSec)]);
end
indClear=cellfun(@(x) ischar(x) || isnan(x),chSec);
nameSec(indClear)=[];
chSec(indClear)=[];
%% �����洢
[~,indSort]=sort(chIdy(:,1),'descend');
chIdyR=cat(2,nameIdy(indSort),num2cell(chIdy(indSort,:)));
xlswrite(fileName,chIdyR,'��ҵ�Ƿ�');

[~,indSort]=sort(peIdy(:,3),'descend');
peIdyR=cat(2,nameIdy(indSort),num2cell(peIdy(indSort,:)));
xlswrite(fileName,peIdyR,'��ҵ��ֵ');

[~,indSort]=sort(cell2mat(chSec),'descend');
chSecR=[nameSec,chSec];
chSecR=chSecR(indSort([1:10,end-9:end]),:);
xlswrite(fileName,chSecR,'�����Ƿ�');