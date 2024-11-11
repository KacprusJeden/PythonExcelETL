DROP TABLE IF EXISTS queuexlsx.QueueDefinitions;
CREATE TABLE queuexlsx.QueueDefinitions (
  id bigint,
  name varchar(100),
  maxnumberofretries int,
  acceptautomaticallyretry int,
  tenantid int,
  isdeleted int,
  deleteruserid bigint,
  deletiontime timestamp,
  lastmodificationtime timestamp,
  lastmodifieruserid bigint,
  creationtime timestamp,
  creatoruserid bigint,
  organizationunitid bigint,
  enforceuniquereference int,
  slainminutes int,
  riskslainminutes int,
  releaseid bigint
);

