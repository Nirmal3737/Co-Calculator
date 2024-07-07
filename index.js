var express = require("express")
var bodyParser = require("body-parser")
var mongoose = require("mongoose")
const multer = require('multer');
const path = require('path');
const xlsx = require('xlsx');
const { spawn } = require('child_process')
const fs = require('fs');
const util = require('util');
const spawnPromise = util.promisify(spawn);

const { Int32 } = require("mongodb")
var nirm = "";
var crrntusr="";
var crrntcode="";
var crrntyear="";
var asstype ="";
var ms ="";
var ls="";
var tot = 0;
var stdattd = 0;
var student = 0;
var uptype="";
var tar =0;
var maxma = 0;
var qcount=0;
const app = express()

app.use(bodyParser.json())
app.use(express.static('public'))
app.use(bodyParser.urlencoded({
    extended:true
}))

mongoose.connect('mongodb+srv://21i339:NirmalKumar@cluster0.rbrfyc2.mongodb.net/?retryWrites=true&w=majority',{
    useNewUrlParser: true,
    useUnifiedTopology: true
});

var db = mongoose.connection;

db.on('error',()=>console.log("Error in Connecting to Database"));
db.once('open',()=>console.log("Connected to Database"))
app.post("/signup",async(req,res)=>{
    try{
        var Name = req.body.name;
        var password = req.body.password;
        var usertype = req.body.details_login;
        var check;
        var data={
            Name : Name,
            Password: password
        }
        if (usertype === 'hod') {
            check = await db.collection('hod').insertOne(data);
        } else if (usertype === 'faculty') {
            check = await db.collection('faculty').insertOne(data);
        }
        return res.redirect("login.html");
    }
    catch (error) {
        console.error('Error during Signup:', error);
        res.send("Error during Signup");}
})
app.get("/question/:courseId/:years",async(req,res)=>{
    try{
        const courseId = req.params.courseId;
        var years = req.params.years;
        var hs = courseId +'_'+ crrntusr;
        console.log(hs);
        var check = await db.collection(hs).findOne({years:years});
        var data={
            cos : check.cos
        }
        for (let i = 1; i <=parseInt( check.cos); i++) {
            const questionKey = `co` + i;
            console.log(questionKey);
            data[questionKey] = check[questionKey];
        }
        console.log("data",data);
        res.json({ success:true, data: data });
    }
    catch(error)
    {
        res.json({ success: false, message: 'Error' });

    }
})
app.get('/editData/:courseId/:years', async (req, res) => {
    try {
        // Fetch data from MongoDB
        const courseId = req.params.courseId;
        var years = req.params.years;
        var hs = courseId + '_'+crrntusr;
        var check = await db.collection(hs).findOne({years:years})
        var data ={
            Batch : check.batch,
            Year : check.year,
            Code : check.code,
            Title : check.title,
            Credit : check.credit,
            Type : check.type,
            Years : check.years,
            Sem: check.sem,
            Class : check.class,
            CO : check.cos,
            poarray  :check.poarray
        }
        for (let i = 1; i <=parseInt( check.cos); i++) {
            const questionKey = `co` + i;
            console.log(questionKey);
            data[questionKey] = check[questionKey];
        }
        res.json({ success:true, data: data });

    } catch (error) {
        console.error(error);
        res.status(500).send('Internal Server Error');
    }
});
app.post("/login",async(req,res)=>{
    try {
        console.log('entered');
        var Name = req.body.name;
        var password = req.body.password;
        var usertype = req.body.details_login;
        console.log(usertype);
        // Use findOne to retrieve a document based on the name
        var check;
        
        if (usertype === 'hod') {
            check = await db.collection('hod').findOne({ Name: Name });
        } else if (usertype === 'faculty') {
            check = await db.collection('faculty').findOne({ Name: Name });
        }
        console.log(check);
        if (check) {
            console.log(check);
            console.log(check.Password);

            // Compare the passwords
            if (check.Password === password) {
                crrntusr= check.Name;
                console.log(crrntusr);
                if(usertype==='faculty')
                    return res.redirect('home.html');
                else if(usertype==='hod')
                    return res.redirect('hodhome.html');
            } else {
                res.send("Wrong password");
            }
        } else {
            res.send("User not found");
        }
    } catch (error) {
        console.error('Error during login:', error);
        res.send("Error during login");}
});
app.post('/process_form', async(req, res) => {
    const code = req.body.code;
    const years = req.body.years;
    const className = req.body.class;
    if(className!="Both")
    {
            
            var check = await db.collection(code).find({years:years, class:className}).toArray()

            const [year, season] = years.split('-');

    // Convert year to a number and calculate the previous year
            const prevYear = parseInt(year) - 1;
            const prevSeason = parseInt(season)-1;

    // Create the output string by concatenating the previous year and the season
            const output = `${prevYear}-${prevSeason}`;
            var check1 = await db.collection(code).find({years:output}).toArray()

            let accumulatedValues = null;
            let count = 0;

            // Iterate over each document
            for (const document of check1) {
                const averageArray = document.codata.map(value => parseFloat(value));
                count++;
                console.log(averageArray);
                if (accumulatedValues === null) {
                    accumulatedValues = averageArray;
                } else {
                    // Add the values of the current document to the accumulated values
                    accumulatedValues = accumulatedValues.map((value, index) => {
                        const currentValue = averageArray[index];
                        return (currentValue !== undefined) ? value + currentValue : value;
                    });
                }
            }

            // Check if accumulatedValues is not null before calculating the average
            const averageResult = (accumulatedValues !== null) ? accumulatedValues.map(value => value / count) : null;

            // Print the resulting average if it exists
            if (averageResult !== null) {
                console.log("Average of the average arrays across documents:", averageResult);
            } else {
                console.log("No data found for the specified filter.");
            }
            const responseData = {
                code:code,
                year:check[0].years,
                codata:check[0].codata,
                codata1:averageResult,
                year1:check1[0].years,
            };

            res.json(responseData);
    }
    else{
        var check = await db.collection(code).find({years:years,class:"IT-G1"}).toArray()
        var check2 = await db.collection(code).find({years:years,class:"IT-G2"}).toArray()

        const [year, season] = years.split('-');

    // Convert year to a number and calculate the previous year
        const prevYear = parseInt(year) - 1;
        const prevSeason = parseInt(season)-1;

    // Create the output string by concatenating the previous year and the season
        const output = `${prevYear}-${prevSeason}`;
            var check1 = await db.collection(code).find({years:output}).toArray()

            var accumulatedValues = null;
            var count = 0;

            // Iterate over each document
            for (const document of check1) {
                const averageArray = document.codata.map(value => parseFloat(value));
                count++;
                console.log(averageArray);
                if (accumulatedValues === null) {
                    accumulatedValues = averageArray;
                } else {
                    // Add the values of the current document to the accumulated values
                    accumulatedValues = accumulatedValues.map((value, index) => {
                        const currentValue = averageArray[index];
                        return (currentValue !== undefined) ? value + currentValue : value;
                    });
                }
            }

            // Check if accumulatedValues is not null before calculating the average
            const averageResult = (accumulatedValues !== null) ? accumulatedValues.map(value => value / count) : null;

            // Print the resulting average if it exists
            if (averageResult !== null) {
                console.log("Average of the average arrays across documents:", averageResult);
            } else {
                console.log("No data found for the specified filter.");
            }
            var accumulatedValues = null;
            var count = 0;

            // Iterate over each document
            for (const document of check) {
                const averageArray1 = document.codata.map(value => parseFloat(value));
                count++;
                console.log(averageArray1);
                if (accumulatedValues === null) {
                    accumulatedValues = averageArray1;
                } else {
                    // Add the values of the current document to the accumulated values
                    accumulatedValues = accumulatedValues.map((value, index) => {
                        const currentValue = averageArray1[index];
                        return (currentValue !== undefined) ? value + currentValue : value;
                    });
                }
            }

            // Check if accumulatedValues is not null before calculating the average
            const averageResult1 = (accumulatedValues !== null) ? accumulatedValues.map(value => value / count) : null;

            // Print the resulting average if it exists
            if (averageResult1 !== null) {
                console.log("Average of the average arrays across documents:", averageResult1);
            } else {
                console.log("No data found for the specified filter.");
            }
            const responseData = {
                code:code,
                year:check[0].years,
                codata:averageResult1,
                codata1:averageResult,
                year1:check1[0].years,
                other:check2[0].codata,
                other1:check[0].codata
            };

            res.json(responseData);
        
    }
});
app.post("/upload",async(req,res)=>{
        var data = {
            "batch": req.body.Batch,
            "year": req.body.Year,
            "code": req.body.Code,
            "title": req.body.Title,
            "credit": req.body.Credit,
            "type": req.body.Type,
            "years": req.body.Years,
            "sem": req.body.Sem,
            "class": req.body.Class,
            "cos": req.body.CO
        }
        console.log(data);
        var coData = {};
        var posarray =[];
        for (var i = 1; i <=req.body.CO; i++) { // Start from 1 to skip the header row
            var rowValues = []; // Array to store values of current row
    
            // Loop through each cell (textbox) in the row
            for (var j = 1; j <= 12; j++) { // Assuming selectedNumber is the number of COs
                var textboxId = "CO" + i + "PO" + j;
                console.log(textboxId); // Constructing the ID of the textbox
                var textboxValue = req.body[textboxId]// Getting the value of the textbox
                rowValues.push(textboxValue); // Pushing the value into the rowValues array
            }
            for(var k =1;k<=2;k++)
            {
                var textboxId = "CO" + i + "PSO" + k;
                console.log(textboxId); // Constructing the ID of the textbox
                var textboxValue = req.body[textboxId];
                console.log(textboxValue); // Getting the value of the textbox
                rowValues.push(textboxValue); 
            }
            posarray.push(rowValues); // Pushing the rowValues array into the main valuesArray
        }
        console.log(posarray);
        // Iterate over the CO textboxes and store their content in the coData object
        for (let i = 1; i <= req.body.CO; i++) {
            coData[`co${i}`] = req.body[`CO${i}`];
        }

        // Combine the main data and CO data
        var documentData = {
            "code": req.body.Code,
            "faculty":crrntusr,
            "year": req.body.Year,
            "years": req.body.Years,
            "class":  req.body.Class,
            "title": req.body.Title,
            "credit": req.body.Credit,
            "type": req.body.Type,
            "poarray": posarray,
            ...coData
        };
        var docData = {
            ...data,
            "poarray":posarray,
            ...coData
        }
        if(req.body.Code!=null && req.body.Year!=null && req.body.Years!=null && req.body.Class!=null && req.body.Title!=null && req.body.Credit!=null && req.body.Type!=null && posarray!=null && coData!=null)
        {
    db.collection(crrntusr).insertOne(data,(err,collection)=>{
        if(err){
            throw err;
        }
        console.log("rINSERED");
    });
}    
            try{
                console.log("Code:",req.body.Code);
                ms = req.body.Code +"_"+crrntusr;
                ls = req.body.Code +"_"+crrntusr+"_type";
             hs = req.body.Code+"_"+crrntusr+"_studentwise";
                await db.createCollection(req.body.Code);
            await db.createCollection(ms);
            await db.createCollection(ls);
            await db.createCollection(hs);


            await db.collection(req.body.Code).insertOne(documentData);
            await db.collection(ms).insertOne(docData);
            console.log("Collection created and data inserted successfully");
            }
            catch(error)
            {
                console.log("Already exists");
            }
    return res.redirect('cocalculator.html');

});
app.post("/postoreData",async(req,res)=>{
    var newData = req.body.newData;
    var code = req.body.code;
    console.log("new",newData);
    var data={
        posarray: newData
    }
    var hs= code +'_'+crrntusr;
    var updatedResult = await db.collection(hs).updateOne(
        { code: code, years: crrntyear }, // Match criteria
        { $set: data } // Update with the new data
        );
})
app.post('/storeData', async(req, res) => {
    var newData = req.body.newData;
    var data ={
        codata : newData   
    }
    try{
        var ms = crrntcode +"_"+crrntusr;
        var check = await db.collection(ms).findOne({codata : newData});
        if(!check)
        {
            console.log("Hi");
            var updatedResult = await db.collection(ms).updateOne(
                { code: crrntcode, years: crrntyear }, // Match criteria
                { $set: data } // Update with the new data
                );
        }
        else{
            console.log("Hi");

            var updatedResult = await db.collection(ms).updateOne(
                { code: crrntcode, years: crrntyear }, // Match criteria
                { $set: data } // Update with the new data
                );
            
        }
        
        var check = await db.collection(crrntcode).findOne({codata : newData});
        if(!check)
        {
            console.log("Hi");

            var updatedResult = await db.collection(crrntcode).updateOne(
                { code: crrntcode, years: crrntyear, faculty:crrntusr }, // Match criteria
                { $set: data } // Update with the new data
                );
        }
        else{
            console.log("Hi");

            var updatedResult = await db.collection(crrntcode).updateOne(
                { code: crrntcode, years: crrntyear,faculty:crrntusr }, // Match criteria
                { $set: data } // Update with the new data
                );
            
        }
    }
    catch(error)
    {
        console.log("Already exists");
    }
    

  });
// Add this route to your existing Node.js server code
app.delete("/delete/course/:courseId", async (req, res) => {
    try {
        const courseId = req.params.courseId;
        var var1 = crrntusr+"_"+courseId; 
        var var2 = crrntusr+"_"+"type";
        // Assuming 'faculty1' is the collection where data is stored
        const result = await db.collection(crrntusr).deleteOne({ code: courseId });
        await db.collection(var1).drop();
        await db.collection(var2).drop();
        if (result.deletedCount > 0) {
            res.json({ success: true, message: 'Deleted successfully' });
        } else {
            res.json({ success: false, message: 'Course ID not found' });
        }
    } catch (error) {
        console.error('Error during delete:', error);
        res.json({ success: false, message: 'Error during delete' });
    }
});
app.delete("/delete/ass/:assessmentType", async (req, res) => {
    try {
        const assessmentType = req.params.assessmentType;
        ls = crrntcode+"_"+crrntusr+"_"+"type";
        // Assuming 'faculty1' is the collection where data is stored
        const result = await db.collection(ls).deleteOne({ assessmentType:assessmentType });

        if (result.deletedCount > 0) {
            res.json({ success: true, message: 'Deleted successfully' });
        } else {
            res.json({ success: false, message: 'Course ID not found' });
        }
    } catch (error) {
        console.error('Error during delete:', error);
        res.json({ success: false, message: 'Error during delete' });
    }
});
app.get("/calnext", async (req, res) => {
    try {
        // Assuming you want to retrieve data from the 'faculty1' collection
        ls = crrntcode+"_"+crrntusr+"_"+"type";

        console.log(ls);
        const subjectData = await db.collection(ls).find().toArray();

        // Extract the relevant fields from the documents
        
        var varData = subjectData.map(doc => {
            if (doc.assessmentType === "Final_sem") { // Check if assessmentType is "ca"
                return {
                    assessmentType: doc.assessmentType,
                    studentsAttended: doc.studentsAttended,
                    target: doc.target,
                    maxMarks: doc.maxMarks,
                    uploadType: doc.uploadType // Include the uploadType field if assessmentType is "ca"
                };
            } else {
                return {
                    assessmentType: doc.assessmentType,
                    studentsAttended: doc.studentsAttended,
                    target: doc.target,
                    maxMarks: doc.maxMarks
                };
            }
        });
        
    
        
        console.log(varData);
        // Send the formatted data as a response
        res.json({ data: varData });
    } catch (error) {
        console.error('Error during fetch:', error);
        res.status(500).send("Error during fetch");
    }
});
app.get("/analysis",async(req,res)=>{
    try{
        var checkdata = await db.collection(crrntcode).find().toArray();
        console.log(checkdata);
        console.log(checkdata[0].credit);
        console.log(checkdata[0].codata);
        var data ={
            code:crrntcode,
            codata : checkdata[0].codata
        }
        console.log(data);
        res.json({data:data});
    }
    catch(error)
    {
        console.error('Error during fetch:', error);
        res.status(500).send("Error during fetch");
    }
})
app.get("/analysis1",async(req,res)=>{
    try{
        var checkdata = await db.collection(crrntcode).find({years:crrntyear}).toArray();
        const [year, season] = crrntyear.split('-');

    // Convert year to a number and calculate the previous year
    const prevYear = parseInt(year) - 1;
    const prevSeason = parseInt(season)-1;

    // Create the output string by concatenating the previous year and the season
    const output = `${prevYear}-${prevSeason}`;
        var checkdata1 = await db.collection(crrntcode).find({years:output}).toArray();
        console.log(checkdata1);
        let accumulatedValues = null;
let count = 0;

// Iterate over each document
for (const document of checkdata1) {
    const averageArray = document.codata.map(value => parseFloat(value));
    count++;
    console.log(averageArray);
    if (accumulatedValues === null) {
        accumulatedValues = averageArray;
    } else {
        // Add the values of the current document to the accumulated values
        accumulatedValues = accumulatedValues.map((value, index) => {
            const currentValue = averageArray[index];
            return (currentValue !== undefined) ? value + currentValue : value;
        });
    }
}

// Check if accumulatedValues is not null before calculating the average
const averageResult = (accumulatedValues !== null) ? accumulatedValues.map(value => value / count) : null;

// Print the resulting average if it exists
if (averageResult !== null) {
    console.log("Average of the average arrays across documents:", averageResult);
} else {
    console.log("No data found for the specified filter.");
}
    if(checkdata1)
    {
        var data ={
            code: crrntcode,
            codata : checkdata[0].codata,
            codata1 : averageResult,
            year1 : checkdata1[0].years,
            year: checkdata[0].years
        }
    }
    else{
        var data ={
            code: crrntcode,
            codata : checkdata[0].codata,
          
            year: checkdata[0].years
        }
    }
        console.log(data);
        res.json({data:data});
    }
    catch(error)
    {
        console.error('Error during fetch:', error);
        res.status(500).send("Error during fetch");
    }
})
// Add this route to your existing Node.js server code
app.get("/cal", async (req, res) => {
    try {
        const subjectData = await db.collection(crrntusr).find().toArray();

        // Extract the relevant fields from the documents
        const formattedData = subjectData.map(doc => ({
            semester: doc.sem,
            years: doc.years,
            credits: doc.credit,
            courseId: doc.code,
            batch: doc.batch,
            year: doc.year,
            title: doc.title,
            type: doc.type,
            cos: doc.cos
        }));

        // Send the formatted data as a response
        res.json({ data: formattedData });
    } catch (error) {
        console.error('Error during fetch:', error);
        res.status(500).send("Error during fetch");
    }
});
app.get("/final/first/:courseId", async (req, res) => {
    try {
        const courseId = req.params.courseId;
        console.log("+new",courseId,"+new");
        crrntcode = courseId;
        // Assuming you want to retrieve data from the 'faculty1' collection
        const subjectData = await db.collection(crrntusr).findOne({ code: courseId });
        console.log(subjectData);
        // If the course ID is found, send the data as a response
        if (subjectData) {
            const formattedData = {
                Semester: subjectData.sem,
                Academic_Year: subjectData.years,
                Credits: subjectData.credit,
                Course_Id: subjectData.code,
                Batch: subjectData.batch,
                Year_of_Study: subjectData.year,
                Title: subjectData.title,
                Type: subjectData.type,
                COs: subjectData.cos
            };
            console.log("sending data");
            res.json({ success:true, data: formattedData });
        } else {
            console.log("not sent");
            res.json({ success: false, message: 'Course ID not found' });
        }
    } catch (error) {
        console.error('Error during fetch:', error);
        res.status(500).send("Error during fetch");
    }
});
app.get("/final/second/:assessmentType", async (req, res) => {
    try {
        const assessmentType = req.params.assessmentType;
        console.log("+new",assessmentType,"+new");
        asstype = assessmentType;
        ls=crrntcode+'_'+crrntusr+'_'+'type';        // Assuming you want to retrieve data from the 'faculty1' collection
        const subjectData = await db.collection(ls).findOne({ assessmentType:assessmentType });
        console.log(subjectData);
        // If the course ID is found, send the data as a response
        if (subjectData) {
            const formattedData = {
                
                assessmentType: subjectData.assessmentType,
                studentsAttended: subjectData.studentsAttended,
                target: subjectData.target,
                maxMarks:subjectData.maxMarks,
                questions: subjectData.questionCount
            };
            for (let i = 1; i <= subjectData.questionCount; i++) {
                const questionKey = `question${i}_answer`;
                const answerKey = `question${i}_co`;
                formattedData[questionKey] = subjectData[questionKey];
                formattedData[answerKey] = subjectData[answerKey];
            }

            console.log("sending data",formattedData);
            res.json({ success:true, data: formattedData });
        } else {
            console.log("not sent");
            res.json({ success: false, message: 'Course ID not found' });
        }
    } catch (error) {
        console.error('Error during fetch:', error);
        res.status(500).send("Error during fetch");
    }
});
app.get("/view2/:courseId", async (req, res) => {
    try {
        const courseId = req.params.courseId;
        console.log("+",courseId,"+");
        crrntcode = courseId;
        
        var hs = crrntcode+"_"+crrntusr+"_studentwise";
        // Assuming you want to retrieve data from the 'faculty1' collection
        const subjectData = await db.collection(hs).find().toArray();
        console.log("view2",subjectData);
        // If the course ID is found, send the data as a response
        if (subjectData) {
            console.log("Sending data");
            console.log(subjectData);
            res.json({ success:true, data: subjectData });
        } else {
            console.log("not sent");
            res.json({ success: false, message: 'Course ID not found' });
        }
    } catch (error) {
        console.error('Error during fetch:', error);
        res.status(500).send("Error during fetch");
    }
});
app.get("/view1/:courseId", async (req, res) => {
    try {
        const courseId = req.params.courseId;
        console.log("+",courseId,"+");
        crrntcode = courseId;
        ls=crrntcode+'_'+crrntusr+'_'+'type';        
        // Assuming you want to retrieve data from the 'faculty1' collection
        const subjectData = await db.collection(ls).find().toArray();
        console.log(subjectData);
        // If the course ID is found, send the data as a response
        if (subjectData) {
            console.log("Sending data");
            console.log(subjectData);
            res.json({ success:true, data: subjectData });
        } else {
            console.log("not sent");
            res.json({ success: false, message: 'Course ID not found' });
        }
    } catch (error) {
        console.error('Error during fetch:', error);
        res.status(500).send("Error during fetch");
    }
});
app.get("/poview/:courseId/:years", async (req, res) => {
    try {
        const courseId = req.params.courseId;
        var years = req.params.years;
        console.log("sddf",years);
        console.log("+",courseId,"+");
        crrntcode = courseId;
        crrntyear = years;
        var ls = crrntcode+'_'+crrntusr;
        // Assuming you want to retrieve data from the 'faculty1' collection
        const subjectData = await db.collection(ls).findOne({ code: courseId,years:years });
        console.log(subjectData);
        // If the course ID is found, send the data as a response
        if (subjectData) {
            const formattedData = {
                podata: subjectData.poarray,
                poatt: subjectData.poattainment
            };
            console.log("sending data");
            res.json({ success:true, data: formattedData });
        } else {
            console.log("not sent");
            res.json({ success: false, message: 'Course ID not found' });
        }
    } catch (error) {
        console.error('Error during fetch:', error);
        res.status(500).send("Error during fetch");
    }
});
app.get("/view/:courseId/:years", async (req, res) => {
    try {
        const courseId = req.params.courseId;
        var years = req.params.years;
        console.log("sddf",years);
        console.log("+",courseId,"+");
        crrntcode = courseId;
        crrntyear = years;
        // Assuming you want to retrieve data from the 'faculty1' collection
        const subjectData = await db.collection(crrntusr).findOne({ code: courseId });
        console.log(subjectData);
        // If the course ID is found, send the data as a response
        if (subjectData) {
            const formattedData = {
                semester: subjectData.sem,
                years: subjectData.years,
                credits: subjectData.credit,
                courseId: subjectData.code,
                batch: subjectData.batch,
                year: subjectData.year,
                title: subjectData.title,
                type: subjectData.type,
                cos: subjectData.cos
            };
            tot_co =parseInt(subjectData.cos);
            var hs = crrntcode+'_'+crrntusr;
            var check = await db.collection(hs).findOne({code:crrntcode,years:crrntyear});
            console.log('ch',check);
            if(check.codata)
            {
                console.log("Available");
                var array = check.poarray;
                var codata = check.codata;

                var cos = parseInt(check.cos);
                var index=0;
                var po=[]
                var cosum = 0
                for(let i=0;i<cos;i++)
                {
                    cosum = cosum + parseFloat(codata[i]);
                }
                var coavg = cosum/cos;
                for(let i=0;i<=13;i++)
                {
                    var sum=0;
                    var colsum=0;
                    for(let j=0;j<cos;j++)
                    {

                        if(array[j][i]!='0')
                        {

                            colsum = colsum + parseFloat(array[j][i]);
                            sum++;
                        }
                    }
                    var poavg = colsum/sum;
                    po[index] = (poavg*coavg)/3;
                    index++;
                }
                console.log(po);
                var data={
                    poattainment: po
                }
                var updateResult = await db.collection(hs).updateOne(
                    { code:crrntcode, years:crrntyear}, // Match criteria
                    { $set: data} // Update with the new data
                    );    
            }
            else{
                console.log("not avail")
            }
            console.log("sending data");
            res.json({ success:true, data: formattedData });
        } else {
            console.log("not sent");
            res.json({ success: false, message: 'Course ID not found' });
        }
    } catch (error) {
        console.error('Error during fetch:', error);
        res.status(500).send("Error during fetch");
    }
});
app.get("/coview/:courseId/:years", async (req, res) => {
    try {
        const courseId = req.params.courseId;
         crrntyear = req.params.years;
        console.log(crrntyear)
        console.log("+",courseId,"+");
        crrntcode = courseId;
        var hs= crrntcode + "_"+crrntusr;
        // Assuming you want to retrieve data from the 'faculty1' collection
        const subjectData = await db.collection(hs).findOne({ code: courseId,years:crrntyear });
        console.log(subjectData);
        var formattedData = {

        };
        // If the course ID is found, send the data as a response
        if (subjectData) {
            for (let i = 1; i <= parseInt(subjectData.cos); i++) {
                const questionKey = `co${i}`;
                formattedData[questionKey] = subjectData[questionKey];
                
            }
            console.log("sending data");
            res.json({ success:true, data: formattedData });
        } else {
            console.log("not sent");
            res.json({ success: false, message: 'Course ID not found' });
        }
    } catch (error) {
        console.error('Error during fetch:', error);
        res.status(500).send("Error during fetch");
    }
});
app.get("/calnext/:courseId/:years", async (req, res) => {
    try {
        const courseId = req.params.courseId;
         crrntyear = req.params.years;
        console.log(crrntyear)
        console.log("+",courseId,"+");
        crrntcode = courseId;
        // Assuming you want to retrieve data from the 'faculty1' collection
        const subjectData = await db.collection(crrntusr).findOne({ code: courseId });
        console.log(subjectData);
        // If the course ID is found, send the data as a response
        if (subjectData) {
            const formattedData = {
                Semester: subjectData.sem,
                Academic_Year: subjectData.years,
                Credits: subjectData.credit,
                Course_Id: subjectData.code,
                Batch: subjectData.batch,
                Year_of_Study: subjectData.year,
                Title: subjectData.title,
                Type: subjectData.type,
                COs: subjectData.cos
            };
            console.log("sending data");
            res.json({ success:true, data: formattedData });
        } else {
            console.log("not sent");
            res.json({ success: false, message: 'Course ID not found' });
        }
    } catch (error) {
        console.error('Error during fetch:', error);
        res.status(500).send("Error during fetch");
    }
});
app.post('/finalfeedback', async(req, res) => {
    // Extract the farray from the request body
    const farray = req.body.farray;
    var data ={
         farray:farray
    }
    var ls = crrntcode+'_'+crrntusr+'_'+'type'; 
    var updateResult = await db.collection(ls).updateOne(
        { assessmentType: "feedback" }, // Match criteria
        { $set: data} // Update with the new data
        );    
    console.log('Received farray:', farray);

    // Send a response back to the client
    res.json({ success: true, message: 'farray received successfully' });
});
app.post("/calnext", async (req, res) => {
    try {
        // Extract data from the request body
        const assessmentType = req.body.as_type;
        const studentsAttended = req.body.Stud_pres;
        const target = req.body.Target;
        const maxMarks = req.body.Tot_marks;
        if(assessmentType=="Final_sem")
        {
            var uploadtype = req.body.finalSemOption;
        }
        asstype = assessmentType;
        var questionCount = req.body.questionCount;
        if(assessmentType=="Final_sem")
        {
            console.log("yes")
            if(uploadtype=="percent")
            {
                console.log("yes")
                uptype=uploadtype;
                var classpercent = req.body.classPercentage;
                var dataToInsert ={
                    assessmentType,
                    studentsAttended,
                    target,
                    maxMarks,
                    questionCount,
                    uploadtype,
                    classpercent
                }
                
            }
            else{
                var dataToInsert ={
                    assessmentType,
                    studentsAttended,
                    target,
                    maxMarks,
                    questionCount,
                    uploadtype
                }
            }
        }
        else{

        
        var dataToInsert = {
            assessmentType,
            studentsAttended,
            target,
            maxMarks,
            questionCount
        };
        }
        for (let i = 1; i <= questionCount; i++) {
            const questionKey = `question${i}_answer`;
            const answerKey = `question${i}_co`;
            
            dataToInsert[questionKey] = req.body[`answer${i}`][0];
            dataToInsert[answerKey] = req.body[`coCheckbox${i}`];
        }
        console.log(dataToInsert);
        ls = crrntcode+"_"+crrntusr+"_"+"type";
        if(assessmentType!=null)
        {
        await db.collection(ls).insertOne(dataToInsert);
        }
        console.log("ass",asstype);
        console.log(uptype);
        if(asstype=="Final_sem" && uptype=="percent")
        {
        console.log("Entered");
        var classpercent = parseInt(req.body.classPercentage);
        var subjectData1 = await db.collection(crrntusr).findOne({ code: crrntcode });
        var tot_co= parseInt(subjectData1.cos);
        var co_arr = new Array(tot_co).fill(0);
        for(let i=1;i<=tot_co;i++)
        {
            co_arr[i-1]="CO" + i;
        }
        var val = new Array(tot_co).fill(classpercent);
        console.log("co arr",co_arr);
        var att = new Array(tot_co).fill(0);
        for(let i=0;i<tot_co;i++)
        {
            if(val[i]>=0 && val[i]<=60)
           {
               att[i]=1;
               att[i]=att[i].toFixed(2);
           }
           else if(val[i]>=61 && val[i]<=79)
           {
               att[i]=2.00;
               att[i]=att[i].toFixed(2);

           }
           else if(val[i]>=80)
           {
               att[i]=3.00;
               att[i]=att[i].toFixed(2);

           }
        }
        console.log("attt",att);
        var dataToUpdate = {};
        for (let i = 0; i < tot_co; i++) {
        dataToUpdate[co_arr[i]] = att[i];
        }
        console.log(dataToUpdate);
        var updateResult = await db.collection(ls).updateOne(
        { assessmentType: asstype }, // Match criteria
        { $set: dataToUpdate} // Update with the new data
        );
            var check = await db.collection(ls).findOne({assessmentType:'CA 1'})
            var std = check.studentsAttended;
            var datas = check.data;
            console.log(datas);
            for(let j=0;j<std;j++)
            {
                
                var gs = crrntcode+'_'+crrntusr+'_'+datas[j][0];
                console.log(gs);
                for (let i = 0; i < co_arr.length; i++) {
                    dataToUpdate[co_arr[i]] = att[i];
                 }
                 var finaldata={
                    "Assessment_type":asstype,
                    ...dataToUpdate
                 };
                 var data={
                    ...dataToUpdate
                 }
                 try{
                    await db.createCollection(gs);
                    var check = await db.collection(gs).findOne({ Assessment_type:asstype });
                    console.log(check);
                    if(check)
                    {
                    var updatedResult = await db.collection(gs).updateOne(
                    { Assessment_type: asstype }, // Match criteria
                    { $set: data } // Update with the new data
                    );
                    console.log("SUCCESSFULLY UPDATED")
                    }
                    else
                    {
                    await db.collection(gs).insertOne(finaldata);
                    console.log("SUCCESS INSERTED");
                    }
                    
                 }
                 catch(error){
                    var check = await db.collection(gs).findOne({ Assessment_type:asstype });
                    console.log(check);
                    if(check)
                    {
                    var updatedResult = await db.collection(gs).updateOne(
                    { Assessment_type: asstype }, // Match criteria
                    { $set: data } // Update with the new data
                    );
                    console.log("SUCCESSFULLY UPDATED")
                    }
                    else
                    {
                    await db.collection(gs).insertOne(finaldata);
                    console.log("SUCCESS INSERTED");
                    }
                 }
                 
                        
                    }
            }
            console.log(crrntcode);
            console.log(crrntyear);
            
            return res.redirect(`/calnext.html?courseId=${crrntcode}&years=${crrntyear}`);
    
}catch (error) {
        console.error('Error during data insertion:', error);
        return res.status(500).send("Error during data insertion");
    }
});
app.post('/storeFirstTableData', async (req, res) => {
    try {
        // Extract firstTableData from request body
        const { firstTableData } = req.body;
        // Check if firstTableData is defined
        if (!firstTableData) {
            throw new Error('firstTableData is undefined');
        }
        const documents = firstTableData.map(data => ({
            _id: new mongoose.Types.ObjectId(), // Generate a new ObjectId for each document
            ...data
        }));

        // Store firstTableData in the database collection
        // Replace `YourModel` with your actual Mongoose model
        await db.collection(ls).insertMany(documents);


        // Send a success response back to the client
        res.status(200).json({ success: true, message: 'Data stored successfully' });
    } catch (error) {
        console.error('Error storing data:', error);
        res.status(500).json({ success: false, message: 'Internal server error' });
    }
});
app.post("/passenger",async(req,res)=>{

    var data = {
        "name" : req.body.name,
        "email" : req.body.email,
        phone : parseInt(req.body.phone),
        seats : parseInt(req.body.seats)
    }
    if(tot>=parseInt(req.body.seats))
    {
        db.collection(nirm).insertOne(data,(err,collection)=>{
            if(err){
                throw err;
            }
            console.log("Record Inserted Successfully");
        });
        return res.redirect('success.html');
    }
    else{
        return res.send("No Seats available");
    }
})
app.post("/sign_up",async(req,res)=>{
    var name = req.body.name;
    var password = req.body.password;

    var data = {
        "name": name,
        "password" : password
    }

    db.collection('users').insertOne(data,(err,collection)=>{
        if(err){
            throw err;
        }
        console.log("Record Inserted Successfully");
    });

    return res.redirect('signup_success.html')

});
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        // Set the destination folder for uploads
        cb(null, 'uploads/');
    },
    filename: (req, file, cb) => {
        // Set the file name to be the original name
        cb(null, file.originalname);
    },
});

const upload = multer({ storage: storage });

// Serve your static files (assuming you have an 'uploads' folder)
app.use('/uploads', express.static('uploads'));

async function convertAndHandle(pdfPath, excelPath, res) {
    try {
        const result = await new Promise((resolve, reject) => {
            convert_pdf(pdfPath, excelPath, (error, result) => {
                if (error) {
                    reject(error);
                } else {
                    resolve(result);
                }
            });
        });
    
       console.log('Conversion result:', result);
       console.log("Conversion Completed");
       // Rest of your code...
   var allSheetData =[];
   const workbook = xlsx.readFile(excelPath);
   const sheetNames = workbook.SheetNames;
   for (let sheetIndex = 0; sheetIndex < sheetNames.length; sheetIndex++) {
       const sheetName = sheetNames[sheetIndex];
       const sheet = workbook.Sheets[sheetName];

       // Convert the sheet to JSON
       const sheetData = xlsx.utils.sheet_to_json(sheet, { header: 1 });

       // Store or process the sheet data accordingly
       allSheetData.push(...sheetData);
   }
   fs.unlink(excelPath, (err) => {
    if (err) {
        console.error('Error deleting file:', err);
        return;
    }
    console.log('Excel file deleted successfully');
});
fs.unlink(pdfPath, (err) => {
    if (err) {
        console.error('Error deleting file:', err);
        return;
    }
    console.log('Excel file deleted successfully');
});
       
       console.log("len",allSheetData.length);
       console.log(allSheetData);
       console.log("Converted success")
       ls=crrntcode+'_'+crrntusr+'_'+'type';        // Assuming you want to retrieve data from the 'faculty1' collection
       const subjectData = await db.collection(ls).findOne({ assessmentType:asstype });
       stdattd=subjectData.studentsAttended;
       tar=parseInt(subjectData.target);
       tar = tar/100;
       console.log(tar);
       maxma=subjectData.maxMarks;
       qcount= subjectData.questionCount;
       var filteredSheetData = [];
       var newArrayOfNumbers=[];
       var uploadtypes = "markpdf";
       var rowdel=0;
       if(asstype=="Final_sem")
       {
        var uploadtypes = subjectData.uploadtype;
       }
       if (
        asstype !== "Assgn" &&
        asstype !== "Tt 1" &&
        asstype !== "Tt 2" &&
        (uploadtypes === "markpdf" || uploadtypes === "gradepdf") &&
        asstype !== "IR 1" &&
        asstype !== "IR 2" &&
        asstype !== "PLR 1" &&
        asstype !== "PLR 1" &&
        asstype !== "Final Lab"
    )       {
           for (let i = 0; i < allSheetData.length; i++)
            {
               var row = allSheetData[i];
               if (isNumber(row[0]) && row[0]!='') {
                    var newRow =row.slice(1);
                    newRow = removeEmptyElements(newRow);
                   newArrayOfNumbers.push(newRow);
               }
               if (typeof row[0] === 'string' && row[0].indexOf('\n') !== -1)
               {
                    var splitted = row[0].split('\n')
                    if(isNumber(splitted[0])&&splitted[1].substring(0, 2) === '21'|| splitted[1].substring(0, 2) === '22'|| splitted[1].substring(0, 2) === '20')
                    {
                        row[0] = splitted[1]
                        row = removeEmptyElements(row);
                        newArrayOfNumbers.push(row);
                    }
                    if(isNumber(splitted[1])&&splitted[0].substring(0, 2) === '21'|| splitted[0].substring(0, 2) === '22'|| splitted[0].substring(0, 2) === '20')
                    {
                        row[0] = splitted[0]
                        row = removeEmptyElements(row);
                        newArrayOfNumbers.push(row);
                    }
               }
           }
       
           for (let i = 0; i < newArrayOfNumbers.length; i++) {
               const row = newArrayOfNumbers[i];
               var check = 0;
               for (let j = 2; j <= 1+parseInt(qcount); j++) {
                   if (row[j] != 0) {
                       check++;
                   }
               }
               if (check!=0) {
                   filteredSheetData.push(row);
               }
               else{
                   rowdel++;
               }
           }
       }
       else{
            console.log("ene");
           for (let i = 0; i < allSheetData.length; i++) {
               var row = allSheetData[i];
               const firstIndex = row[0];
               if(row && row[0] )
               {
                   if (typeof firstIndex === 'string' && firstIndex.substring(0, 2) === '21'|| firstIndex.substring(0, 2) === '22'|| firstIndex.substring(0, 2) === '20' ) {
                    row=removeEmptyElements(row);   
                    newArrayOfNumbers.push(row);
                   }
               }
               
           }
           filteredSheetData.push(...newArrayOfNumbers);
       }
// Print the new array of rows with a number in the first index

console.log("fi",filteredSheetData);
   console.log(rowdel);
       console.log(qcount);
       var columnSums = new Array(parseInt(qcount)).fill(0);
       console.log("s",columnSums.length);
       qcount = parseInt(qcount);
       var qs=1;
       student  = stdattd;
       var tot_std =stdattd;
       stdattd=stdattd-rowdel;
       console.log("asdhb",stdattd);
       if(asstype!="Assgn" && asstype!="Tt 1" && asstype!="Tt 2" && uploadtypes=="markpdf" && asstype!="IR 1" && asstype!="IR 2"&& asstype!="PLR 1"&& asstype!="PLR 1"&& asstype!="Final Lab")
       {
        console.log("Entered CA");
           for (let i = 2; i <= 2+qcount-1; i++) {
               var c=0;
               var marksKey = `question${qs}_answer`;
               var max1 = parseInt(subjectData[marksKey]);
               console.log(max1);
               var thresmark = tar * max1;
               console.log(thresmark);
               for(let j=0;j<stdattd;j++)
               {
                   if(parseFloat(filteredSheetData[j][i])>thresmark)
                   {
                       c++;
                   }
               }
               columnSums[qs-1]=c;
               qs++;
           }
           
           for(let i=0;i<stdattd;i++)
           {
            qs=1;
           var columnSums1 = filteredSheetData[i].slice(2,2+qcount);
           console.log(columnSums1);
           var columnper1 = new Array(qcount).fill(0);
                    for(let j=0;j<qcount;j++)
                    {
                        var marksKey = `question${qs}_answer`;
                        var max1 = parseInt(subjectData[marksKey]);
                        console.log("Max",max1);
                        columnper1[j]=(columnSums1[j]/max1)*100;
                        qs++;
                    }
                        var coq1 = new Array(qcount).fill(0);
                        for (let j = 1; j <= qcount; j++) {
                            var answerKey = `question${j}_co`;
                            coq1[j-1] = subjectData[answerKey];
                        }

                    var coq = [];
                    var columnper = [];

                    // Iterate through array 'a'
                    for (let j = 0; j < coq1.length; j++) {
                    // If the element is an array, flatten it
                    if (Array.isArray(coq1[j])) {
                        coq = coq.concat(coq1[j]);
                        // Duplicate elements in 'b' for each element in the nested array
                        for (let k = 0; k < coq1[j].length; k++) {
                            columnper.push(columnper1[j]);
                        }
                    } else {
                        coq.push(coq1[j]);
                        // If the element is not an array, add 'b[i]' to 'duplicatedB' once
                        columnper.push(columnper1[j]);
                    }
                    }
                    console.log("Sushanth",columnper);
                    var att = new Array(coq.length).fill(0);
       
                    for(let j=0;j<columnper.length;j++)
                    {
                        if(columnper[j]>=0 && columnper[j]<=60)
                        {
                            att[j]=1;
                        }
                        else if(columnper[j]>=61 && columnper[j]<=79)
                        {
                            att[j]=2;
                        }
                        else if(columnper[j]>=80)
                        {
                            att[j]=3;
                        }
                    }
                    console.log("qtt arra",att);
                    
                        var coMap = new Map();
                    
                        // Iterate over the co and marks arrays
                        for (let j = 0; j < coq.length; j++) {
                            var coElement = coq[j];
                            var marksElement = att[j];
                    
                            // If the coElement is not in the map, initialize it with an array [marksElement, 1]
                            if (!coMap.has(coElement)) {
                                coMap.set(coElement, [marksElement, 1]);
                            } else {
                                // If coElement is already in the map, update the sum and count
                                var [sum, count] = coMap.get(coElement);
                                coMap.set(coElement, [sum + marksElement, count + 1]);
                            }
                        }
                    
                        // Calculate the average for each coElement
                        var coArray = [];
                        var averageMarksArray = [];
                    
                        for (const [coElement, [sum, count]] of coMap) {
                            var averageMarks = (sum / count).toFixed(2);
                            coArray.push(coElement);
                            averageMarksArray.push(averageMarks);
                        }
                        console.log(coArray);
                        console.log(averageMarksArray);
                        var dataToUpdate = {};
                        var gs= crrntcode+"_"+crrntusr+"_"+filteredSheetData[i][0];
             for (let i = 0; i < coArray.length; i++) {
                dataToUpdate[coArray[i]] = averageMarksArray[i];
             }
             var finaldata={
                "Assessment_type":asstype,
                ...dataToUpdate
             };
             var data={
                ...dataToUpdate
             }
             try{
                await db.createCollection(gs);
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESSFULLY CREATED");
                
             }
             catch(error){
                var check = await db.collection(gs).findOne({ Assessment_type:asstype });
                console.log(check);
                if(check)
                {
                var updatedResult = await db.collection(gs).updateOne(
                { Assessment_type: asstype }, // Match criteria
                { $set: data } // Update with the new data
                );
                console.log("SUCCESSFULLY UPDATED")
                }
                else
                {
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESS INSERTED");
                }
             }
           }
       }
       if(asstype=="Final_sem" && uploadtypes=="gradepdf")
       {
        console.log("Final Semester ");
        var c=0;
        var columnSums = new Array(1).fill(0);

        var thresmark = tar * 10;        
        for(let j=0;j<stdattd;j++)
        {
            if(parseFloat(filteredSheetData[j][2])>thresmark)
            {
                c++;
            }
        }
        columnSums[qs-1]=c;
        qs++;
        
       }
       qs = 1;
       if(asstype=="Assgn")
       {
               console.log("Entered Assignment Presentation");
               var c=0;
               var marksKey = `question${qs}_answer`;
               var max1 = parseInt(subjectData[marksKey]);
               console.log(max1);
               var thresmark = tar * max1;
               console.log(thresmark);
               for(let j=0;j<stdattd;j++)
               {
                   console.log(filteredSheetData[j][6]);
                   if(parseFloat(filteredSheetData[j][6])>thresmark)
                   {
                       c++;
                   }
               }
               columnSums[qs-1]=c;
               qs++;
               //student wise
               for(let i=0;i<stdattd;i++)
               {
                    var columnSums1 = filteredSheetData[i][6];
                    var columnper1 = new Array(qcount).fill(0);
                    qs=1;
                    for(let j=0;j<qcount;j++)
                    {
                        var marksKey = `question${qs}_answer`;
                    var max1 = parseInt(subjectData[marksKey]);
                    console.log("Max",max1);
                        columnper1[j]=(columnSums1/max1)*100;
                        qs++;
                    }
                        var coq1 = new Array(qcount).fill(0);
                        for (let j = 1; j <= qcount; j++) {
                            var answerKey = `question${j}_co`;
                            coq1[j-1] = subjectData[answerKey];
                        }

                    var coq = [];
                    var columnper = [];

                    // Iterate through array 'a'
                    for (let j = 0; j < coq1.length; j++) {
                    // If the element is an array, flatten it
                    if (Array.isArray(coq1[j])) {
                        coq = coq.concat(coq1[j]);
                        // Duplicate elements in 'b' for each element in the nested array
                        for (let k = 0; k < coq1[j].length; k++) {
                            columnper.push(columnper1[j]);
                        }
                    } else {
                        coq.push(coq1[j]);
                        // If the element is not an array, add 'b[i]' to 'duplicatedB' once
                        columnper.push(columnper1[j]);
                    }
                    }
                    console.log("Sushanth",columnper);
                    var att = new Array(coq.length).fill(0);
       
                    for(let j=0;j<columnper.length;j++)
                    {
                        if(columnper[j]>=0 && columnper[j]<=60)
                        {
                            att[j]=1;
                        }
                        else if(columnper[j]>=61 && columnper[j]<=79)
                        {
                            att[j]=2;
                        }
                        else if(columnper[j]>=80)
                        {
                            att[j]=3;
                        }
                    }
                    console.log("qtt arra",att);
                    
                        var coMap = new Map();
                    
                        // Iterate over the co and marks arrays
                        for (let j = 0; j < coq.length; j++) {
                            var coElement = coq[j];
                            var marksElement = att[j];
                    
                            // If the coElement is not in the map, initialize it with an array [marksElement, 1]
                            if (!coMap.has(coElement)) {
                                coMap.set(coElement, [marksElement, 1]);
                            } else {
                                // If coElement is already in the map, update the sum and count
                                var [sum, count] = coMap.get(coElement);
                                coMap.set(coElement, [sum + marksElement, count + 1]);
                            }
                        }
                    
                        // Calculate the average for each coElement
                        var coArray = [];
                        var averageMarksArray = [];
                    
                        for (const [coElement, [sum, count]] of coMap) {
                            var averageMarks = (sum / count).toFixed(2);
                            coArray.push(coElement);
                            averageMarksArray.push(averageMarks);
                        }
                        console.log(coArray);
                        console.log(averageMarksArray);
                        var dataToUpdate = {};
                        var gs= crrntcode+"_"+crrntusr+"_"+filteredSheetData[i][0];
             for (let i = 0; i < coArray.length; i++) {
                dataToUpdate[coArray[i]] = averageMarksArray[i];
             }
             var finaldata={
                "Assessment_type":asstype,
                ...dataToUpdate
             };
             var data={
                ...dataToUpdate
             }
             try{
                await db.createCollection(gs);
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESSFULLY CREATED");
                
             }
             catch(error){
                var check = await db.collection(gs).findOne({ Assessment_type:asstype });
                console.log(check);
                if(check)
                {
                var updatedResult = await db.collection(gs).updateOne(
                { Assessment_type: asstype }, // Match criteria
                { $set: data } // Update with the new data
                );
                console.log("SUCCESSFULLY UPDATED")
                }
                else
                {
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESS INSERTED");
                }
             }
             
                    
                }
       }
       qs = 1;
       if(asstype=="IR 1")
       {
               console.log("Entered IR1");
               var c=0;
               var marksKey = `question${qs}_answer`;
               var max1 = parseInt(subjectData[marksKey]);
               console.log(max1);
               var thresmark = tar * max1;
               console.log(thresmark);
               for(let j=0;j<stdattd;j++)
               {
                   console.log(filteredSheetData[j][2]);
                   if(parseFloat(filteredSheetData[j][2])>thresmark)
                   {
                       c++;
                   }
               }
               columnSums[qs-1]=c;
               qs++;
               //student wise
               for(let i=0;i<stdattd;i++)
               {
                    var columnSums1 = filteredSheetData[i][2];
                    var columnper1 = new Array(qcount).fill(0);
                    qs=1;
                    for(let j=0;j<qcount;j++)
                    {
                        var marksKey = `question${qs}_answer`;
                    var max1 = parseInt(subjectData[marksKey]);
                    console.log("Max",max1);
                        columnper1[j]=(columnSums1/max1)*100;
                        qs++;
                    }
                        var coq1 = new Array(qcount).fill(0);
                        for (let j = 1; j <= qcount; j++) {
                            var answerKey = `question${j}_co`;
                            coq1[j-1] = subjectData[answerKey];
                        }

                    var coq = [];
                    var columnper = [];

                    // Iterate through array 'a'
                    for (let j = 0; j < coq1.length; j++) {
                    // If the element is an array, flatten it
                    if (Array.isArray(coq1[j])) {
                        coq = coq.concat(coq1[j]);
                        // Duplicate elements in 'b' for each element in the nested array
                        for (let k = 0; k < coq1[j].length; k++) {
                            columnper.push(columnper1[j]);
                        }
                    } else {
                        coq.push(coq1[j]);
                        // If the element is not an array, add 'b[i]' to 'duplicatedB' once
                        columnper.push(columnper1[j]);
                    }
                    }
                    console.log("Sushanth",columnper);
                    var att = new Array(coq.length).fill(0);
       
                    for(let j=0;j<columnper.length;j++)
                    {
                        if(columnper[j]>=0 && columnper[j]<=60)
                        {
                            att[j]=1;
                        }
                        else if(columnper[j]>=61 && columnper[j]<=79)
                        {
                            att[j]=2;
                        }
                        else if(columnper[j]>=80)
                        {
                            att[j]=3;
                        }
                    }
                    console.log("qtt arra",att);
                    
                        var coMap = new Map();
                    
                        // Iterate over the co and marks arrays
                        for (let j = 0; j < coq.length; j++) {
                            var coElement = coq[j];
                            var marksElement = att[j];
                    
                            // If the coElement is not in the map, initialize it with an array [marksElement, 1]
                            if (!coMap.has(coElement)) {
                                coMap.set(coElement, [marksElement, 1]);
                            } else {
                                // If coElement is already in the map, update the sum and count
                                var [sum, count] = coMap.get(coElement);
                                coMap.set(coElement, [sum + marksElement, count + 1]);
                            }
                        }
                    
                        // Calculate the average for each coElement
                        var coArray = [];
                        var averageMarksArray = [];
                    
                        for (const [coElement, [sum, count]] of coMap) {
                            var averageMarks = (sum / count).toFixed(2);
                            coArray.push(coElement);
                            averageMarksArray.push(averageMarks);
                        }
                        console.log(coArray);
                        console.log(averageMarksArray);
                        var dataToUpdate = {};
                        var gs= crrntcode+"_"+crrntusr+"_"+filteredSheetData[i][0];
             for (let i = 0; i < coArray.length; i++) {
                dataToUpdate[coArray[i]] = averageMarksArray[i];
             }
             var finaldata={
                "Assessment_type":asstype,
                ...dataToUpdate
             };
             var data={
                ...dataToUpdate
             }
             try{
                await db.createCollection(gs);
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESSFULLY CREATED");
                
             }
             catch(error){
                var check = await db.collection(gs).findOne({ Assessment_type:asstype });
                console.log(check);
                if(check)
                {
                var updatedResult = await db.collection(gs).updateOne(
                { Assessment_type: asstype }, // Match criteria
                { $set: data } // Update with the new data
                );
                console.log("SUCCESSFULLY UPDATED")
                }
                else
                {
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESS INSERTED");
                }
             }
             
                    
                }
       }
       qs=1;
       if(asstype=="Final Lab")
       {
               console.log("fi lab");
               var c=0;
               var marksKey = `question${qs}_answer`;
               var max1 = parseInt(subjectData[marksKey]);
               console.log(max1);
               var thresmark = tar * max1;
               console.log(thresmark);
               for(let j=0;j<stdattd;j++)
               {
                   console.log(filteredSheetData[j][4]);
                   if(parseFloat(filteredSheetData[j][4])>thresmark)
                   {
                       c++;
                   }
               }
               columnSums[qs-1]=c;
               qs++;
               //student wise
               for(let i=0;i<stdattd;i++)
               {
                    var columnSums1 = filteredSheetData[i][4];
                    var columnper1 = new Array(qcount).fill(0);
                    qs=1;
                    for(let j=0;j<qcount;j++)
                    {
                        var marksKey = `question${qs}_answer`;
                    var max1 = parseInt(subjectData[marksKey]);
                    console.log("Max",max1);
                        columnper1[j]=(columnSums1/max1)*100;
                        qs++;
                    }
                        var coq1 = new Array(qcount).fill(0);
                        for (let j = 1; j <= qcount; j++) {
                            var answerKey = `question${j}_co`;
                            coq1[j-1] = subjectData[answerKey];
                        }

                    var coq = [];
                    var columnper = [];

                    // Iterate through array 'a'
                    for (let j = 0; j < coq1.length; j++) {
                    // If the element is an array, flatten it
                    if (Array.isArray(coq1[j])) {
                        coq = coq.concat(coq1[j]);
                        // Duplicate elements in 'b' for each element in the nested array
                        for (let k = 0; k < coq1[j].length; k++) {
                            columnper.push(columnper1[j]);
                        }
                    } else {
                        coq.push(coq1[j]);
                        // If the element is not an array, add 'b[i]' to 'duplicatedB' once
                        columnper.push(columnper1[j]);
                    }
                    }
                    console.log("Sushanth",columnper);
                    var att = new Array(coq.length).fill(0);
       
                    for(let j=0;j<columnper.length;j++)
                    {
                        if(columnper[j]>=0 && columnper[j]<=60)
                        {
                            att[j]=1;
                        }
                        else if(columnper[j]>=61 && columnper[j]<=79)
                        {
                            att[j]=2;
                        }
                        else if(columnper[j]>=80)
                        {
                            att[j]=3;
                        }
                    }
                    console.log("qtt arra",att);
                    
                        var coMap = new Map();
                    
                        // Iterate over the co and marks arrays
                        for (let j = 0; j < coq.length; j++) {
                            var coElement = coq[j];
                            var marksElement = att[j];
                    
                            // If the coElement is not in the map, initialize it with an array [marksElement, 1]
                            if (!coMap.has(coElement)) {
                                coMap.set(coElement, [marksElement, 1]);
                            } else {
                                // If coElement is already in the map, update the sum and count
                                var [sum, count] = coMap.get(coElement);
                                coMap.set(coElement, [sum + marksElement, count + 1]);
                            }
                        }
                    
                        // Calculate the average for each coElement
                        var coArray = [];
                        var averageMarksArray = [];
                    
                        for (const [coElement, [sum, count]] of coMap) {
                            var averageMarks = (sum / count).toFixed(2);
                            coArray.push(coElement);
                            averageMarksArray.push(averageMarks);
                        }
                        console.log(coArray);
                        console.log(averageMarksArray);
                        var dataToUpdate = {};
                        var gs= crrntcode+"_"+crrntusr+"_"+filteredSheetData[i][0];
             for (let i = 0; i < coArray.length; i++) {
                dataToUpdate[coArray[i]] = averageMarksArray[i];
             }
             var finaldata={
                "Assessment_type":asstype,
                ...dataToUpdate
             };
             var data={
                ...dataToUpdate
             }
             try{
                await db.createCollection(gs);
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESSFULLY CREATED");
                
             }
             catch(error){
                var check = await db.collection(gs).findOne({ Assessment_type:asstype });
                console.log(check);
                if(check)
                {
                var updatedResult = await db.collection(gs).updateOne(
                { Assessment_type: asstype }, // Match criteria
                { $set: data } // Update with the new data
                );
                console.log("SUCCESSFULLY UPDATED")
                }
                else
                {
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESS INSERTED");
                }
             }
             
                    
                }
       }
       qs = 1;
       if(asstype=="IR 2")
       {
               console.log("Entered Assignment Presentation");
               var c=0;
               var marksKey = `question${qs}_answer`;
               var max1 = parseInt(subjectData[marksKey]);
               console.log(max1);
               var thresmark = tar * max1;
               console.log(thresmark);
               for(let j=0;j<stdattd;j++)
               {
                   console.log(filteredSheetData[j][3]);
                   if(parseFloat(filteredSheetData[j][3])>thresmark)
                   {
                       c++;
                   }
               }
               columnSums[qs-1]=c;
               qs++;
               //student wise
               for(let i=0;i<stdattd;i++)
               {
                    var columnSums1 = filteredSheetData[i][3];
                    var columnper1 = new Array(qcount).fill(0);
                    qs=1;
                    for(let j=0;j<qcount;j++)
                    {
                        var marksKey = `question${qs}_answer`;
                    var max1 = parseInt(subjectData[marksKey]);
                    console.log("Max",max1);
                        columnper1[j]=(columnSums1/max1)*100;
                        qs++;
                    }
                        var coq1 = new Array(qcount).fill(0);
                        for (let j = 1; j <= qcount; j++) {
                            var answerKey = `question${j}_co`;
                            coq1[j-1] = subjectData[answerKey];
                        }

                    var coq = [];
                    var columnper = [];

                    // Iterate through array 'a'
                    for (let j = 0; j < coq1.length; j++) {
                    // If the element is an array, flatten it
                    if (Array.isArray(coq1[j])) {
                        coq = coq.concat(coq1[j]);
                        // Duplicate elements in 'b' for each element in the nested array
                        for (let k = 0; k < coq1[j].length; k++) {
                            columnper.push(columnper1[j]);
                        }
                    } else {
                        coq.push(coq1[j]);
                        // If the element is not an array, add 'b[i]' to 'duplicatedB' once
                        columnper.push(columnper1[j]);
                    }
                    }
                    console.log("Sushanth",columnper);
                    var att = new Array(coq.length).fill(0);
       
                    for(let j=0;j<columnper.length;j++)
                    {
                        if(columnper[j]>=0 && columnper[j]<=60)
                        {
                            att[j]=1;
                        }
                        else if(columnper[j]>=61 && columnper[j]<=79)
                        {
                            att[j]=2;
                        }
                        else if(columnper[j]>=80)
                        {
                            att[j]=3;
                        }
                    }
                    console.log("qtt arra",att);
                    
                        var coMap = new Map();
                    
                        // Iterate over the co and marks arrays
                        for (let j = 0; j < coq.length; j++) {
                            var coElement = coq[j];
                            var marksElement = att[j];
                    
                            // If the coElement is not in the map, initialize it with an array [marksElement, 1]
                            if (!coMap.has(coElement)) {
                                coMap.set(coElement, [marksElement, 1]);
                            } else {
                                // If coElement is already in the map, update the sum and count
                                var [sum, count] = coMap.get(coElement);
                                coMap.set(coElement, [sum + marksElement, count + 1]);
                            }
                        }
                    
                        // Calculate the average for each coElement
                        var coArray = [];
                        var averageMarksArray = [];
                    
                        for (const [coElement, [sum, count]] of coMap) {
                            var averageMarks = (sum / count).toFixed(2);
                            coArray.push(coElement);
                            averageMarksArray.push(averageMarks);
                        }
                        console.log(coArray);
                        console.log(averageMarksArray);
                        var dataToUpdate = {};
                        var gs= crrntcode+"_"+crrntusr+"_"+filteredSheetData[i][0];
             for (let i = 0; i < coArray.length; i++) {
                dataToUpdate[coArray[i]] = averageMarksArray[i];
             }
             var finaldata={
                "Assessment_type":asstype,
                ...dataToUpdate
             };
             var data={
                ...dataToUpdate
             }
             try{
                await db.createCollection(gs);
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESSFULLY CREATED");
                
             }
             catch(error){
                var check = await db.collection(gs).findOne({ Assessment_type:asstype });
                console.log(check);
                if(check)
                {
                var updatedResult = await db.collection(gs).updateOne(
                { Assessment_type: asstype }, // Match criteria
                { $set: data } // Update with the new data
                );
                console.log("SUCCESSFULLY UPDATED")
                }
                else
                {
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESS INSERTED");
                }
             }
             
                    
                }
       }
       qs = 1;
       if(asstype=="PLR 1")
       {
               console.log("Entered Assignment Presentation");
               var c=0;
               var marksKey = `question${qs}_answer`;
               var max1 = parseInt(subjectData[marksKey]);
               console.log(max1);
               var thresmark = tar * max1;
               console.log(thresmark);
               for(let j=0;j<stdattd;j++)
               {
                   console.log(filteredSheetData[j][4]);
                   if(parseFloat(filteredSheetData[j][4])>thresmark)
                   {
                       c++;
                   }
               }
               columnSums[qs-1]=c;
               qs++;
               //student wise
               for(let i=0;i<stdattd;i++)
               {
                    var columnSums1 = filteredSheetData[i][4];
                    var columnper1 = new Array(qcount).fill(0);
                    qs=1;
                    for(let j=0;j<qcount;j++)
                    {
                        var marksKey = `question${qs}_answer`;
                    var max1 = parseInt(subjectData[marksKey]);
                    console.log("Max",max1);
                        columnper1[j]=(columnSums1/max1)*100;
                        qs++;
                    }
                        var coq1 = new Array(qcount).fill(0);
                        for (let j = 1; j <= qcount; j++) {
                            var answerKey = `question${j}_co`;
                            coq1[j-1] = subjectData[answerKey];
                        }

                    var coq = [];
                    var columnper = [];

                    // Iterate through array 'a'
                    for (let j = 0; j < coq1.length; j++) {
                    // If the element is an array, flatten it
                    if (Array.isArray(coq1[j])) {
                        coq = coq.concat(coq1[j]);
                        // Duplicate elements in 'b' for each element in the nested array
                        for (let k = 0; k < coq1[j].length; k++) {
                            columnper.push(columnper1[j]);
                        }
                    } else {
                        coq.push(coq1[j]);
                        // If the element is not an array, add 'b[i]' to 'duplicatedB' once
                        columnper.push(columnper1[j]);
                    }
                    }
                    console.log("Sushanth",columnper);
                    var att = new Array(coq.length).fill(0);
       
                    for(let j=0;j<columnper.length;j++)
                    {
                        if(columnper[j]>=0 && columnper[j]<=60)
                        {
                            att[j]=1;
                        }
                        else if(columnper[j]>=61 && columnper[j]<=79)
                        {
                            att[j]=2;
                        }
                        else if(columnper[j]>=80)
                        {
                            att[j]=3;
                        }
                    }
                    console.log("qtt arra",att);
                    
                        var coMap = new Map();
                    
                        // Iterate over the co and marks arrays
                        for (let j = 0; j < coq.length; j++) {
                            var coElement = coq[j];
                            var marksElement = att[j];
                    
                            // If the coElement is not in the map, initialize it with an array [marksElement, 1]
                            if (!coMap.has(coElement)) {
                                coMap.set(coElement, [marksElement, 1]);
                            } else {
                                // If coElement is already in the map, update the sum and count
                                var [sum, count] = coMap.get(coElement);
                                coMap.set(coElement, [sum + marksElement, count + 1]);
                            }
                        }
                    
                        // Calculate the average for each coElement
                        var coArray = [];
                        var averageMarksArray = [];
                    
                        for (const [coElement, [sum, count]] of coMap) {
                            var averageMarks = (sum / count).toFixed(2);
                            coArray.push(coElement);
                            averageMarksArray.push(averageMarks);
                        }
                        console.log(coArray);
                        console.log(averageMarksArray);
                        var dataToUpdate = {};
                        var gs= crrntcode+"_"+crrntusr+"_"+filteredSheetData[i][0];
             for (let i = 0; i < coArray.length; i++) {
                dataToUpdate[coArray[i]] = averageMarksArray[i];
             }
             var finaldata={
                "Assessment_type":asstype,
                ...dataToUpdate
             };
             var data={
                ...dataToUpdate
             }
             try{
                await db.createCollection(gs);
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESSFULLY CREATED");
                
             }
             catch(error){
                var check = await db.collection(gs).findOne({ Assessment_type:asstype });
                console.log(check);
                if(check)
                {
                var updatedResult = await db.collection(gs).updateOne(
                { Assessment_type: asstype }, // Match criteria
                { $set: data } // Update with the new data
                );
                console.log("SUCCESSFULLY UPDATED")
                }
                else
                {
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESS INSERTED");
                }
             }
             
                    
                }
       }
       qs=1;
       if(asstype=="PLR 2")
       {
               console.log("Entered Assignment Presentation");
               var c=0;
               var marksKey = `question${qs}_answer`;
               var max1 = parseInt(subjectData[marksKey]);
               console.log(max1);
               var thresmark = tar * max1;
               console.log(thresmark);
               for(let j=0;j<stdattd;j++)
               {
                   console.log(filteredSheetData[j][5]);
                   if(parseFloat(filteredSheetData[j][5])>thresmark)
                   {
                       c++;
                   }
               }
               columnSums[qs-1]=c;
               qs++;
               //student wise
               for(let i=0;i<stdattd;i++)
               {
                    var columnSums1 = filteredSheetData[i][5];
                    var columnper1 = new Array(qcount).fill(0);
                    qs=1;
                    for(let j=0;j<qcount;j++)
                    {
                        var marksKey = `question${qs}_answer`;
                    var max1 = parseInt(subjectData[marksKey]);
                    console.log("Max",max1);
                        columnper1[j]=(columnSums1/max1)*100;
                        qs++;
                    }
                        var coq1 = new Array(qcount).fill(0);
                        for (let j = 1; j <= qcount; j++) {
                            var answerKey = `question${j}_co`;
                            coq1[j-1] = subjectData[answerKey];
                        }

                    var coq = [];
                    var columnper = [];

                    // Iterate through array 'a'
                    for (let j = 0; j < coq1.length; j++) {
                    // If the element is an array, flatten it
                    if (Array.isArray(coq1[j])) {
                        coq = coq.concat(coq1[j]);
                        // Duplicate elements in 'b' for each element in the nested array
                        for (let k = 0; k < coq1[j].length; k++) {
                            columnper.push(columnper1[j]);
                        }
                    } else {
                        coq.push(coq1[j]);
                        // If the element is not an array, add 'b[i]' to 'duplicatedB' once
                        columnper.push(columnper1[j]);
                    }
                    }
                    console.log("Sushanth",columnper);
                    var att = new Array(coq.length).fill(0);
       
                    for(let j=0;j<columnper.length;j++)
                    {
                        if(columnper[j]>=0 && columnper[j]<=60)
                        {
                            att[j]=1;
                        }
                        else if(columnper[j]>=61 && columnper[j]<=79)
                        {
                            att[j]=2;
                        }
                        else if(columnper[j]>=80)
                        {
                            att[j]=3;
                        }
                    }
                    console.log("qtt arra",att);
                    
                        var coMap = new Map();
                    
                        // Iterate over the co and marks arrays
                        for (let j = 0; j < coq.length; j++) {
                            var coElement = coq[j];
                            var marksElement = att[j];
                    
                            // If the coElement is not in the map, initialize it with an array [marksElement, 1]
                            if (!coMap.has(coElement)) {
                                coMap.set(coElement, [marksElement, 1]);
                            } else {
                                // If coElement is already in the map, update the sum and count
                                var [sum, count] = coMap.get(coElement);
                                coMap.set(coElement, [sum + marksElement, count + 1]);
                            }
                        }
                    
                        // Calculate the average for each coElement
                        var coArray = [];
                        var averageMarksArray = [];
                    
                        for (const [coElement, [sum, count]] of coMap) {
                            var averageMarks = (sum / count).toFixed(2);
                            coArray.push(coElement);
                            averageMarksArray.push(averageMarks);
                        }
                        console.log(coArray);
                        console.log(averageMarksArray);
                        var dataToUpdate = {};
                        var gs= crrntcode+"_"+crrntusr+"_"+filteredSheetData[i][0];
             for (let i = 0; i < coArray.length; i++) {
                dataToUpdate[coArray[i]] = averageMarksArray[i];
             }
             var finaldata={
                "Assessment_type":asstype,
                ...dataToUpdate
             };
             var data={
                ...dataToUpdate
             }
             try{
                await db.createCollection(gs);
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESSFULLY CREATED");
                
             }
             catch(error){
                var check = await db.collection(gs).findOne({ Assessment_type:asstype });
                console.log(check);
                if(check)
                {
                var updatedResult = await db.collection(gs).updateOne(
                { Assessment_type: asstype }, // Match criteria
                { $set: data } // Update with the new data
                );
                console.log("SUCCESSFULLY UPDATED")
                }
                else
                {
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESS INSERTED");
                }
             }
             
                    
                }
       }
       qs =1;
       if(asstype=="Tt 1")
       {
           console.log("Tutorial1");

               var c=0;
               var marksKey = `question${qs}_answer`;
               var max1 = parseInt(subjectData[marksKey]);
               console.log(max1);
               var thresmark = tar * max1;
               console.log(thresmark);
               for(let j=0;j<stdattd;j++)
               {
                   console.log(filteredSheetData[j][7]);
                   if(parseFloat(filteredSheetData[j][7])>thresmark)
                   {
                       c++;
                   }
               }
               columnSums[qs-1]=c;
               qs++;
               for(let i=0;i<stdattd;i++)
           {
           var columnSums1 = filteredSheetData[i][7];
           var columnper1 = new Array(qcount).fill(0);
           qs=1;
                    for(let j=0;j<qcount;j++)
                    {
                        var marksKey = `question${qs}_answer`;
                        var max1 = parseInt(subjectData[marksKey]);
                        console.log("Max",max1);
                        columnper1[j]=(columnSums1/max1)*100;
                        qs++;
                    }
                        var coq1 = new Array(qcount).fill(0);
                        for (let j = 1; j <= qcount; j++) {
                            var answerKey = `question${j}_co`;
                            coq1[j-1] = subjectData[answerKey];
                        }

                    var coq = [];
                    var columnper = [];

                    // Iterate through array 'a'
                    for (let j = 0; j < coq1.length; j++) {
                    // If the element is an array, flatten it
                    if (Array.isArray(coq1[j])) {
                        coq = coq.concat(coq1[j]);
                        // Duplicate elements in 'b' for each element in the nested array
                        for (let k = 0; k < coq1[j].length; k++) {
                            columnper.push(columnper1[j]);
                        }
                    } else {
                        coq.push(coq1[j]);
                        // If the element is not an array, add 'b[i]' to 'duplicatedB' once
                        columnper.push(columnper1[j]);
                    }
                    }
                    console.log("Sushanth",columnper);
                    var att = new Array(coq.length).fill(0);
       
                    for(let j=0;j<columnper.length;j++)
                    {
                        if(columnper[j]>=0 && columnper[j]<=60)
                        {
                            att[j]=1;
                        }
                        else if(columnper[j]>=61 && columnper[j]<=9)
                        {
                            att[j]=2;
                        }
                        else if(columnper[j]>=80)
                        {
                            att[j]=3;
                        }
                    }
                    console.log("qtt arra",att);
                    
                        var coMap = new Map();
                    
                        // Iterate over the co and marks arrays
                        for (let j = 0; j < coq.length; j++) {
                            var coElement = coq[j];
                            var marksElement = att[j];
                    
                            // If the coElement is not in the map, initialize it with an array [marksElement, 1]
                            if (!coMap.has(coElement)) {
                                coMap.set(coElement, [marksElement, 1]);
                            } else {
                                // If coElement is already in the map, update the sum and count
                                var [sum, count] = coMap.get(coElement);
                                coMap.set(coElement, [sum + marksElement, count + 1]);
                            }
                        }
                    
                        // Calculate the average for each coElement
                        var coArray = [];
                        var averageMarksArray = [];
                    
                        for (const [coElement, [sum, count]] of coMap) {
                            var averageMarks = (sum / count).toFixed(2);
                            coArray.push(coElement);
                            averageMarksArray.push(averageMarks);
                        }
                        console.log(coArray);
                        console.log(averageMarksArray);
                        var dataToUpdate = {};
                        var gs= crrntcode+"_"+crrntusr+"_"+filteredSheetData[i][0];
             for (let i = 0; i < coArray.length; i++) {
                dataToUpdate[coArray[i]] = averageMarksArray[i];
             }
             var finaldata={
                "Assessment_type":asstype,
                ...dataToUpdate
             };
             var data={
                ...dataToUpdate
             }
             try{
                await db.createCollection(gs);
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESSFULLY CREATED");
                
             }
             catch(error){
                var check = await db.collection(gs).findOne({ Assessment_type:asstype });
                console.log(check);
                if(check)
                {
                var updatedResult = await db.collection(gs).updateOne(
                { Assessment_type: asstype }, // Match criteria
                { $set: data } 
                );
                console.log("SUCCESSFULLY UPDATED")
                }
                else
                {
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESS INSERTED");
                }
             }

           }
       }
       qs =1;
       if(asstype=="Tt 2")
       {
           console.log("Tutorial2");

               var c=0;
               var marksKey = `question${qs}_answer`;
               var max1 = parseInt(subjectData[marksKey]);
               console.log(max1);
               var thresmark = tar * max1;
               console.log(thresmark);
               for(let j=0;j<stdattd;j++)
               {
                   console.log(filteredSheetData[j][8]);
                   if(parseFloat(filteredSheetData[j][8])>thresmark)
                   {
                       c++;
                   }
               }
               columnSums[qs-1]=c;
               qs++;
               for(let i=0;i<stdattd;i++)
           {
            qs=1;
           var columnSums1 = filteredSheetData[i][8];
           var columnper1 = new Array(qcount).fill(0);
                    for(let j=0;j<qcount;j++)
                    {
                        var marksKey = `question${qs}_answer`;
                        var max1 = parseInt(subjectData[marksKey]);
                        console.log("Max",max1);
                        columnper1[j]=(columnSums1/max1)*100;
                        qs++;
                    }
                        var coq1 = new Array(qcount).fill(0);
                        for (let j = 1; j <= qcount; j++) {
                            var answerKey = `question${j}_co`;
                            coq1[j-1] = subjectData[answerKey];
                        }

                    var coq = [];
                    var columnper = [];

                    // Iterate through array 'a'
                    for (let j = 0; j < coq1.length; j++) {
                    // If the element is an array, flatten it
                    if (Array.isArray(coq1[j])) {
                        coq = coq.concat(coq1[j]);
                        // Duplicate elements in 'b' for each element in the nested array
                        for (let k = 0; k < coq1[j].length; k++) {
                            columnper.push(columnper1[j]);
                        }
                    } else {
                        coq.push(coq1[j]);
                        // If the element is not an array, add 'b[i]' to 'duplicatedB' once
                        columnper.push(columnper1[j]);
                    }
                    }
                    console.log("Sushanth",columnper);
                    var att = new Array(coq.length).fill(0);
       
                    for(let j=0;j<columnper.length;j++)
                    {
                        if(columnper[j]>=0 && columnper[j]<=60)
                        {
                            att[j]=1;
                        }
                        else if(columnper[j]>=61 && columnper[j]<=79)
                        {
                            att[j]=2;
                        }
                        else if(columnper[j]>=80)
                        {
                            att[j]=3;
                        }
                    }
                    console.log("qtt arra",att);
                    
                        var coMap = new Map();
                    
                        // Iterate over the co and marks arrays
                        for (let j = 0; j < coq.length; j++) {
                            var coElement = coq[j];
                            var marksElement = att[j];
                    
                            // If the coElement is not in the map, initialize it with an array [marksElement, 1]
                            if (!coMap.has(coElement)) {
                                coMap.set(coElement, [marksElement, 1]);
                            } else {
                                // If coElement is already in the map, update the sum and count
                                var [sum, count] = coMap.get(coElement);
                                coMap.set(coElement, [sum + marksElement, count + 1]);
                            }
                        }
                    
                        // Calculate the average for each coElement
                        var coArray = [];
                        var averageMarksArray = [];
                    
                        for (const [coElement, [sum, count]] of coMap) {
                            var averageMarks = (sum / count).toFixed(2);
                            coArray.push(coElement);
                            averageMarksArray.push(averageMarks);
                        }
                        console.log(coArray);
                        console.log(averageMarksArray);
                        var dataToUpdate = {};
                        var gs= crrntcode+"_"+crrntusr+"_"+filteredSheetData[i][0];
             for (let i = 0; i < coArray.length; i++) {
                dataToUpdate[coArray[i]] = averageMarksArray[i];
             }
             var finaldata={
                "Assessment_type":asstype,
                ...dataToUpdate
             };
             var data={
                ...dataToUpdate
             }
             try{
                await db.createCollection(gs);
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESSFULLY CREATED");
                
             }
             catch(error){
                var check = await db.collection(gs).findOne({ Assessment_type:asstype });
                console.log(check);
                if(check)
                {
                var updatedResult = await db.collection(gs).updateOne(
                { Assessment_type: asstype }, // Match criteria
                { $set: data } // Update with the new data
                );
                console.log("SUCCESSFULLY UPDATED")
                }
                else
                {
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESS INSERTED");
                }
             }

           }
               
       }
       var subjectData1 = await db.collection(crrntusr).findOne({ code: crrntcode });
       var tot_co= parseInt(subjectData1.cos);

       for(let i=0;i<tot_std;i++)
       {
            var student = newArrayOfNumbers[i][0];
            var gs=crrntcode+"_"+crrntusr+"_"+student;
            var keyNames = [];
            for(let j=0;j<tot_co;j++)
            {
                keyNames[j] = "CO" + (j+1);
            }
            console.log(gs);
            var sums = {};
            var counts = {};

            await db.collection(gs).find().forEach(function(doc) {
                keyNames.forEach(function(key) {
                    if (key in doc) {
                        // If the key doesn't exist in the sums object, initialize it
                        if (!(key in sums)) {
                            sums[key] = 0;
                            counts[key] = 0;
                        }
                        // Add the value to the sum and increment the count
                        sums[key] += parseFloat(doc[key]);
                        counts[key]++;
                    }
                });
            });

            // Calculate the average for each key
            var averages = {};
            keyNames.forEach(function(key) {
                if (key in sums) {
                    averages[key] = (sums[key] / counts[key]).toFixed(2);
                }
            });
            var finaldata={
                Assessment_type:"overall",
                averages
            }
            var data={
                averages
            }
            var check = await db.collection(gs).findOne({ Assessment_type:"overall" });
                if(check)
                {
                var updatedResult = await db.collection(gs).updateOne(
                { Assessment_type: "overall" }, // Match criteria
                { $set: data } // Update with the new data
                );
                console.log("SUCCESSFULLY UPDATED")
                }
                else
                {
                await db.collection(gs).insertOne(finaldata);
                console.log("SUCCESS INSERTED");
                }
        }
        var hs = crrntcode+"_"+crrntusr+"_studentwise";

        for(let i=0;i<tot_std;i++)
        {
            
                console.log("En");
        
            var student = newArrayOfNumbers[i][0];
            var gs=crrntcode+"_"+crrntusr+"_"+student;
                var check = await db.collection(gs).findOne({Assessment_type:"overall"});
                var data={
                    Roll_no: student,
                    averages: check.averages
                }
                var data1={
                    averages: check.averages
                }
                var check1 = await db.collection(hs).findOne({Roll_no:student});
                if(check1)
                {
                    console.log("Avail")
                    var updatedResult = await db.collection(hs).updateOne(
                        { Roll_no:student }, // Match criteria
                        { $set: data1 } // Update with the new data
                        );
                }

                else
                {
                    console.log("not avail")
                await db.collection(hs).insertOne(data);
                }
            
           
        }
       console.log(columnSums);
       var columnper1 = new Array(qcount).fill(0);
       for(let i=0;i<qcount;i++)
       {
           columnper1[i]=(columnSums[i]/stdattd)*100;
       }
       var coq1 = new Array(qcount).fill(0);
       for (let i = 1; i <= qcount; i++) {
           var answerKey = `question${i}_co`;
           coq1[i-1] = subjectData[answerKey];
       }

var coq = [];
var columnper = [];

// Iterate through array 'a'
for (let i = 0; i < coq1.length; i++) {
   // If the element is an array, flatten it
   if (Array.isArray(coq1[i])) {
       coq = coq.concat(coq1[i]);
       // Duplicate elements in 'b' for each element in the nested array
       for (let j = 0; j < coq1[i].length; j++) {
           columnper.push(columnper1[i]);
       }
   } else {
       coq.push(coq1[i]);
       // If the element is not an array, add 'b[i]' to 'duplicatedB' once
       columnper.push(columnper1[i]);
   }
}
       console.log("hi",qcount);
       var siz = qcount;
       console.log(siz);
       siz = parseInt(siz);
       var att = new Array(coq.length).fill(0);
       
       for(i=0;i<columnper.length;i++)
       {
           if(parseInt(columnper[i])>=0 && parseInt(columnper[i])<=60)
           {
               att[i]=1;
           }
           else if(parseInt(columnper[i])>=61 && parseInt(columnper[i])<=79)
           {
               att[i]=2;
           }
           else if(parseInt(columnper[i])>=80)
           {
               att[i]=3;
           }
       }
       console.log("qtt arra",att);
       
           var coMap = new Map();
       
           // Iterate over the co and marks arrays
           for (let i = 0; i < coq.length; i++) {
               var coElement = coq[i];
               var marksElement = att[i];
       
               // If the coElement is not in the map, initialize it with an array [marksElement, 1]
               if (!coMap.has(coElement)) {
                   coMap.set(coElement, [marksElement, 1]);
               } else {
                   // If coElement is already in the map, update the sum and count
                   var [sum, count] = coMap.get(coElement);
                   coMap.set(coElement, [sum + marksElement, count + 1]);
               }
           }
       
           // Calculate the average for each coElement
           var coArray = [];
           var averageMarksArray = [];
       
           for (const [coElement, [sum, count]] of coMap) {
               var averageMarks = (sum / count).toFixed(2);
               coArray.push(coElement);
               averageMarksArray.push(averageMarks);
           }
           console.log(coArray);
           console.log(averageMarksArray);
           var dataToUpdate = {};
           var colength = coArray.length;
for (let i = 0; i < coArray.length; i++) {
   dataToUpdate[coArray[i]] = averageMarksArray[i];
}
var updateResult = await db.collection(ls).updateOne(
   { assessmentType: asstype }, // Match criteria
   { $set: dataToUpdate } // Update with the new data
);
var updatedResult = await db.collection(ls).updateOne(
   { assessmentType: asstype }, // Match criteria
   { $set: {
       coques: coArray.length
   } } // Update with the new data
);



       console.log(coq);
       // Log or send the sheet data in the rsponse
       console.log(columnper);
       console.log(att);
       ls = crrntcode+"_"+crrntusr+"_"+"type";
       var data={
           "data": newArrayOfNumbers
       }
       db.collection(ls).updateOne({ assessmentType: asstype },
           { $set: data },(err,collection)=>{
           if(err){
               throw err;
           }
           console.log("INSERTED");
       });
       // Send a response to the client
       return res.redirect("cocalculator.html");
    }
       catch (error) {
        console.error('Error converting PDF to Excel:', error);
        throw error; // Re-throw the error for the outer try-catch to handle
    }
}
function removeEmptyElements(arr) {
    // Filter out empty elements from the array
    return arr.filter(element => {
        // Check if the element is not empty
        if (element !== null && element !== undefined && element !== '') {
            // If the element is an object or array, recursively check its elements
            if (Array.isArray(element) || typeof element === 'object') {
                return removeEmptyElements(element).length !== 0;
            }
            return true;
        }
        return false;
    });
}

app.get('/finalpovi/:courseId/:years', async(req, res) => {
    try {
        console.log("Entered");
        const courseId = req.params.courseId;
         crrntyear = req.params.years;
         var hs=crrntcode+'_'+crrntusr;
         var check = await db.collection(hs).findOne({code:crrntcode,years:crrntyear})
         var data={
            poarray:check.poarray
         }
         console.log("poarray",data);
        res.json({ success:true, data: data});

       
    }
    catch(error)
    {
        console.error('Error:', error);
        return res.status(500).json({ success: false, message: 'Internal server error' });
    }
});
app.post('/finalpo',async(req,res)=>{
    var hs=crrntcode+"_"+crrntusr;
    console.log("Enterrrrrr")
    var check = await db.collection(hs).findOne({code:crrntcode,years:crrntyear});
    var cos = parseInt(check.cos);
    var posarray =[];
    for (var i = 1; i <=cos; i++) { // Start from 1 to skip the header row
        var rowValues = []; // Array to store values of current row

        // Loop through each cell (textbox) in the row
        for (var j = 1; j <= 12; j++) { // Assuming selectedNumber is the number of COs
            var textboxId = "CO" + i + "PO" + j;
            console.log(textboxId); // Constructing the ID of the textbox
            var textboxValue = req.body[textboxId]// Getting the value of the textbox
            rowValues.push(textboxValue); // Pushing the value into the rowValues array
        }
        for(var k =1;k<=2;k++)
        {
            var textboxId = "CO" + i + "PSO" + k;
            console.log(textboxId); // Constructing the ID of the textbox
            var textboxValue = req.body[textboxId];
            console.log(textboxValue); // Getting the value of the textbox
            rowValues.push(textboxValue); 
        }
        posarray.push(rowValues); // Pushing the rowValues array into the main valuesArray
    }
    console.log(posarray);
    var data={
        poarray: posarray
    }
    var updatedResult = await db.collection(hs).updateOne(
        { code: crrntcode, years: crrntyear }, // Match criteria
        { $set: data } // Update with the new data
        );
        return res.redirect("cocalculator.html")
})
// Handle file upload
app.post('/final', upload.single('file'), async(req, res) => {
    try {
        const file = req.file;
        ls=crrntcode+'_'+crrntusr+'_'+'type';        // Assuming you want to retrieve data from the 'faculty1' collection

        // Check if a file was uploaded
        if (!file) {
            return res.status(400).json({ success: false, message: 'No file uploaded' });
        }

        // Process the file (you can save it to disk, database, etc.)
        // For demonstration purposes, we're logging the file information
        const fileDetails = {
            originalname: file.originalname,
            encoding: file.encoding,
            mimetype: file.mimetype,
            destination: file.destination,
            filename: file.filename,
            path: file.path,
            size: file.size,
        };
        console.log('File Uploaded',fileDetails);
        var baseName = path.basename(file.originalname, path.extname(file.originalname));
        baseName = ls + '_' + baseName;
        console.log("cv",baseName);
        const pdfPath = file.path;
        const excelPath = path.join('uploads', `${baseName}.xlsx`);
        // Convert the uploaded PDF file to Excel using convertapi
    
       
 await convertAndHandle(pdfPath, excelPath, res);

console.log("COMPLETED");
        
    } catch (error) {
        console.error('Error during file upload:', error);
        return res.status(500).json({ success: false, message: 'Internal server error' });
    }
});
async function convert_pdf(pdfPath, excelPath, callback) {
    const pythonProcess = spawn('python', ['test.py', pdfPath, excelPath]);

    pythonProcess.stdout.on('data', (data) => {
        console.log(`stdout: ${data}`);
    });

    pythonProcess.stderr.on('data', (data) => {
        console.error(`stderr: ${data}`);
    });

    pythonProcess.on('close', (code) => {
        console.log(`child process exited with code ${code}`);
        if (code === 0) {
            console.log('Conversion completed.');
            callback(null, 1); // Pass 1 to indicate success
        } else {
            console.error(`child process exited with code ${code}`);
            callback(new Error(`Conversion failed with code ${code}`)); // Pass an error object
        }
    });
}
function isNumber(value) {
    return !isNaN(value);
}
app.get("/",(req,res)=>{
    res.set({
        "Allow-access-Allow-Origin": '*'
    })
    return res.redirect('login.html');
}).listen(1200);


console.log("Listening on PORT 1200");