create table if not exists runwhat(
    'id' integer primary key not null autoincrement,
    'pack_name' varchar(32) not null,
    'pack_version' varchar(16) not null,
    'pack_env' varchar(16) not null,
    'pack_type' varchar(1) not null default '0',
    'pack_phone_type' varchar(1) not null default '0',
    'pack_create_date' date default NOW(),
    'pack_times' integer default 0
);

create unique index idx_packwhat on packwhat('pack_name','pack_version','pack_env','pack_type');
