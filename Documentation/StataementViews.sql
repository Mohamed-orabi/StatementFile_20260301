


-------------
-------------
-------------
create or replace view V_StatementMasterReward as
select * from tstatementmastertable x 
where x.contracttype = 'Reward Program (Airmile)'
-------------
create or replace view V_StatementDetailReward as
select * from tstatementdetailtable x 
where x.trandescription = 'Calculated Points'


