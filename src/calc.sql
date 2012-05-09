--任务完成情况汇总表
DROP TABLE "task";
CREATE TABLE "task" ("taskgroup" VARCHAR,"taskname" VARCHAR,"endate" DATE,"ET" INTEGER,"actor" VARCHAR,"result" VARCHAR,"quality" VARCHAR,"memo" VARCHAR,"realrate" INTEGER,"realtime" INTEGER,"onerate" INTEGER,"onetime" INTEGER,"tworate" INTEGER,"twotime" INTEGER,"threerate" INTEGER,"threetime" INTEGER,"fourrate" INTEGER,"fourtime" INTEGER,"fiverate" INTEGER,"fivetime" INTEGER,"sixrate" INTEGER,"sixtime" INTEGER,"sevenrate" INTEGER,"seventime" INTEGER,"flag4finish" INTEGER,"flag4delay" INTEGER,"flag4trac" INTEGER,"flag4check" INTEGER)

--任务完成情况统计表
DROP TABLE "calc";
CREATE TABLE "calc" ("taskgroup" VARCHAR,"totalnum" INTEGER,"flag4finish" INTEGER,"flag4result" INTEGER,"flag4delay" INTEGER,"flag4trac" INTEGER,"flag4check" INTEGER,"totalET" float,"totalTT" float)

--计划任务表 
DROP TABLE "taskfromgantproj";
CREATE TABLE "taskfromgantproj" ("id" INTEGER, "name" VARCHAR, "startdate" DATETIME, "duration" INTEGER,"plan" VARCHAR,"strdur" VARCHAR,"pmid" INTEGER)

--资源表
DROP TABLE "resources";
CREATE TABLE "resources" ("id" INTEGER, "name" VARCHAR)

--资源分配表
DROP TABLE "allocations";
CREATE TABLE "allocations" ("taskid" INTEGER, "resourceid" INTEGER)

--任务分配表
DROP TABLE "plan"；
CREATE TABLE "plan" ("plan" VARCHAR, "taskname" VARCHAR, "name" VARCHAR, "enddate" DATETIME, "checker" VARCHAR, "ET" INTEGER)

--根据任务组统计预计工时和实际工时
select taskgroup as '任务组',count(1) as '任务总数',sum(ET)/480 as '累计预计工时',sum(onetime+twotime+threetime+fourtime+fivetime+sixtime+seventime)/480 as '累计实际工时' from task
group by taskgroup

--根据任务组统计完成率
select taskgroup,count(1) from task
where flag4finish = 1
group by taskgroup

--根据任务组统计合格率
select taskgroup,count(1) from task
where result = '合格'
group by taskgroup

--根据任务组统计逾期率
select taskgroup,count(1) from task
where flag4delay = 1
group by taskgroup

--根据任务组统计跟踪率
select taskgroup,count(1) from task
where flag4trac = 1
group by taskgroup

--根据任务组统计检查率
select taskgroup,count(1) from task
where flag4check = 1
group by taskgroup

INSERT INTO calc ("taskgroup","totalnum","totalET","totalTT") select taskgroup as '任务组',count(1) as '任务总数',sum(ET)/480 as '累计预计工时',sum(onetime+twotime+threetime+fourtime+fivetime+sixtime+seventime)/480 as '累计实际工时' from task
group by taskgroup

replace into calc ("flag4finish") select count(1) from task
where flag4finish = 1 group by taskgroup

--统计计划任务
SELECT a.plan,a.name taskname,c.name,date(a.startdate,a.strdur) enddate,(select resources.name from resources where resources.id = a.pmid) FROM taskfromgantproj a,allocations b,resources c
where a.id = b.taskid and b.resourceid = c.id
