function AddAnnouncement(){
  //connect to the spreadsheet and range to get the data
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sheet1");
  //get the first 50 rows from row 2 onwards
  var range = sheet.getRange("A2:E50");
  var values = range.getValues();

  //create a list for each of the Google Classrooms in column A.
  var thisClass = "";
  var classes = [];
  var classAnnouncements = [];
  for (i = 0; i<values.length; i++) {
    thisClass = values[i][0];
    if (thisClass != "") {
      if (classes.includes(thisClass) == false) {
        classes.push(thisClass);
      }  
    }
  
  };
  console.log(classes);

  
  var courses = Classroom.Courses.list().courses; //courses that this teacher has access to 
  var course = courses[1];
  var i = 0;
  var j = 0;
  var foundClass = false;
  const foundClassesTotal = classes.length; //total classes i need to find
  var k = 0; //records how many classes we have found 
  
  //makes a copy of the classes array
  var courseIds = JSON.parse(JSON.stringify(classes));

  //go through the courses this teacher has and for each one, see if it matches the classes requested in 
  //column A. Save it in courseIds array.

  while (i < courses.length && k < foundClassesTotal) {
    course = courses[i];
    foundClass = false;
    //this loop checks if the current course name matches one of our required classes
    j=0;
    while (j < classes.length && foundClass==false) {
      
      if (course.name == classes[j]) {
        foundClass = true;
        courseIds[j] = course.id;
        k++;
      }
      j++;
    }
    
    i++;
    
  }
  
  var proceed = true;
  for (i=0;i<classes.length;i++){
    if (classes[i] == courseIds[i]) {
      SpreadsheetApp.getUi().alert(classes[i] + " is not valid. Please check the class name is correct.");
      proceed = false;
    }
  }
  
  if (proceed) { 
    //all class names are valid, we can create the scheduled announcements in each row of the range

    var COURSEID = "";
    var ClassSource =  {
      text: "STRINGS"+"\n"+"STRINGS",
      
    }; 
    var strDate = new Date(values[0][4]);
    var courseName = values[0][0];
    var timestamp = Utilities.formatDate(strDate, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
    
    var i = 0; //loop counter

  //this loop goes through each row of data and stops when it encounters a row that has a blank for coursename
    while (values[i][0] != "") {
      strDate = new Date(values[i][4]); //get the scheduled date for this row
      timestamp = Utilities.formatDate(strDate, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
      ClassSource = {
        text: values[i][1] + " " + values[i][2] + " " + values[i][3],
        state: "DRAFT",
        scheduledTime: timestamp,
      };
      //COURSEID = getCourseID(values[i][0],courseIds,classes);
      COURSEID = courseIds[classes.indexOf(values[i][0])];

      Classroom.Courses.Announcements.create(ClassSource, COURSEID);
      i++;
    }
  };
  // for (var i = 0; values.length-1; i++) {
  //   strDate = new Date(values[i][1]);
    
  //   timestamp = Utilities.formatDate(strDate, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
    
  //   ClassSource = {
  //     text: values[i][2] + " " + values[i][0] + " " + values[i][3],
  //     state: "DRAFT",
  //     scheduledTime: timestamp,
  //   };
    
  //   Classroom.Courses.Announcements.create(ClassSource, COURSEID)
  // };
  // }
  
  // const CLASS_DATA = Classroom.Courses.list().courses;
  // console.log(CLASS_DATA);
  // const DATA = CLASS_DATA.map(c => {
  //         return [c.name, c.section,c.room,c.description,c.enrollmentCode];
  // });

  

  //set up the classroom announcement structure
     

  

  //set up timestamp structure for use with scheduledtime
  
  //var i = 0; //loop counter

  //this loop goes through each row of data and stops when it encounters a row that has a blank for coursename
  // while (values[i][0] != "") {
  //   strDate = new Date(values[i][4]); //get the scheduled date for this row
  //   timestamp = Utilities.formatDate(strDate, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  //   ClassSource = {
  //     text: values[i][2] + " " + values[i][0] + " " + values[i][3],
  //     state: "DRAFT",
  //     scheduledTime: timestamp,
  //   };

  // }
  // for (var i = 0; values.length-1; i++) {
  //   strDate = new Date(values[i][1]);
    
  //   timestamp = Utilities.formatDate(strDate, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
    
  //   ClassSource = {
  //     text: values[i][2] + " " + values[i][0] + " " + values[i][3],
  //     state: "DRAFT",
  //     scheduledTime: timestamp,
  //   };
    
  //   Classroom.Courses.Announcements.create(ClassSource, COURSEID)
  // };
     
  
  // Logger.log(ClassSource);
  //DATA = [[course.name,course.id,course.room,course.description,course.enrollmentCode]];
  //console.log(DATA);
  


  // const START_ROW = 2;
  // const START_COLUMN = 1
  // sheet.getRange(START_ROW,START_COLUMN,
  //             DATA.length,DATA[0].length).setValues(DATA);

  
}

// function getCourseID(className, idArray, classesArray) {
//   var index = classesArray.indexOf(className);
//   return idArray[index];
// }

