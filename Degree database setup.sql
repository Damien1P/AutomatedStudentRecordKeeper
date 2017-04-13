drop database if exists "studentrecordkeeper";
create database "studentrecordkeeper";
set datestyle to 'SQL, DMY';

create table student
(
	name text,
	studentid integer,
	previouschool text,
	previousprogram	text,
	yearlevel integer,
	yearentered date default Current_date
);
CREATE UNIQUE INDEX studentid ON table (student);

create table courses
(
	coursesubject text ,
	coursenumber integer,
	coursesection text,
	coursename text,
	courseyear integer,
	yearlevel integer,
	credits double precision
);
CREATE UNIQUE INDEX coursename ON table (courses);

create table complementarycourses
(
	coursesubject text,
	coursenumber integer,
	coursename text ,
	credits integer
);
CREATE UNIQUE INDEX coursename ON table (courses);

create table makeupcourses
(
	ID serial primary key,
	studentid integer,
	coursesubject text null constraint valid_coursecode check (length(coursesubject) > 0),
	coursenumber integer,
	coursesection text,
	courseyear integer null constraint valid_year check(courseyear>2000),
	yearlevel integer,
	credits integer
);

create table coursegrades
(
	ID serial primary key,
	coursesubject text null constraint valid_name check (length(name) >0),
	coursenumber integer,
	takenyear integer null constraint valid_year check(courseyear>2000),
	grade integer null constraint valid_grade check(0<=grade<=100)
);



