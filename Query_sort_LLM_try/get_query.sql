select
    '20250720' as tdbank_imp_date,
    t3.relayentity as relayentity,
    t3.deskid as deskid,
    t3.deskseq as deskseq,
    t3.openid as openid,
    t3.chatcontents as chatcontents,
    t3.chattime
from
(
    select
        relaysvrentity, gameseq, gamesvrentity, acntcamp, vopenid
    from ieg_tdbank::smoba_dsl_5v5pvpsettle_fht0
    where 
        substr(tdbank_imp_date, 1, 8) >= '20250720' and 
        CAST(SUBSTRING(uid, LENGTH(uid)-4, 2) AS INT) BETWEEN 80 AND 99
    limit 3000  -- 限制t2查询结果数量以优化性能
) t2
join
(
    select 
        relay_entity as relayentity,
        desk_seq as deskseq,
        desk_id as deskid,
        openid,
        get_json_object(get_json_object(mission, '$.ori_chat_content'), '$.chat_text') as chatcontents,
        frame_no as chattime
    from ieg_tdbank::smoba_ai_dsl_chat_commander_trace_fht0 
    where 
        substr(tdbank_imp_date, 1, 8) >= '20250720' and
        event_type = 'mission' and 
        env = 'formal' and 
        get_json_object(mission, '$.msg_type') = '2'  -- 只获取msg_type=2的数据
    limit 3000  -- 直接在这里限制3000条记录
) t3
on t2.relaysvrentity = t3.relayentity and 
   t2.gameseq = t3.deskseq and 
   t2.gamesvrentity = t3.deskid and 
   t2.vopenid = t3.openid
where t3.chatcontents is not null and t3.chatcontents <> ''
limit 3000  -- 最终结果也限制3000条确保