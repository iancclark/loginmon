.mode csv
.headers on

select username, datetime(start_time,'unixepoch') as start_time, datetime(end_time,'unixepoch') as end_time, (end_time - start_time)/3600.0 as hours, sent FROM sessions order by datetime(start_time);
