# msaccess-jsonFromData
Create JSON from MS Access relational tables, queries and SQL, to any nested depth.
Why? In this case to help test a web service with SOAP UI, so we can edit data for tests in an Access database rather than as JSON.
We use a trick for including properties from one-to-many related tables/queries, involving writing a query that outputs instructions on how to fetch related values from another table or query.
Example query to serialise people and their cats:
SELECT People.*, "GetRelated(Cats,OwnerID,ID)" AS Cats FROM People;
where [Cats] is the related table, Cats.OwnerID is the FK, ID is the key in [People] identifying the owner. This can work to any depth.
To use the code, paste it into a module, then use the public functions:
JSONfromData(sDataSource As String, Optional bPerformValidation As Boolean = False) As String
or 
JSONfromRecordset(rs As Recordset, bPerformValidation As Boolean) As String
The validation verifies that the result is valid JSON, but takes longer. NB it doesn't convert rubbish into JSON.
The JSON validation method was inspired by amadeus' answer on
http://stackoverflow.com/questions/2782076/is-there-a-json-parser-for-vb6-vba
though we are only using that method for validation and prettifying, and not for creating the JSON.
Not all data types are catered for (or tested) - this was just a quickie to get something working for testing with SOAP UI.
