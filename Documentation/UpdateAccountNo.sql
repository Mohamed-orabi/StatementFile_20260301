UPDATE tstatementmastertable e1
SET e1.externalno = (SELECT OLDACCOUNTNO
FROM tacc2acc e2
WHERE e1.accountno = e2.NEWACCOUNTNO and e2.branch = 1)
WHERE e1.branch = 1;