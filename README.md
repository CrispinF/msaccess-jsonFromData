# msaccess-jsonFromData
Create JSON from MS Access relational tables, queries and SQL, to any nested depth.
The JSON serialisation method was initially inspired by amadeus answer on
http://stackoverflow.com/questions/2782076/is-there-a-json-parser-for-vb6-vba
though as it turns out we are only using that method for optional validation and prettifying.
We've added a trick for including properties from one-to-many related tables/queries.
