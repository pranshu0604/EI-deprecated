import React, { useState, useEffect } from 'react'
import { useParams } from 'react-router-dom'
import axios from 'axios'
import { useNavigate } from 'react-router-dom'
import { FaTrash } from "react-icons/fa";
import { IoStatsChart } from "react-icons/io5";
import { IoMdPersonAdd } from "react-icons/io";


const SubDash = () => {
  const { subjectCode } = useParams();
  const [create, setCreate] = useState(false);
  const [schema, setSchema] = useState(false);
  const navigate = useNavigate();

  return (
    <div className="w-full min-h-screen h-full pb-12 poppins">
      {create && <AddStudentPopup setCreate={setCreate} subjectCode={subjectCode} />}
      {schema && <AddExamSchema setSchema={setSchema} subjectCode={subjectCode} />}
      <div>
        <button
          className="bg-gray-200 dark:bg-gray-800 text-black dark:text-white px-4 py-2 rounded"
          onClick={() => navigate(-1)}
        >
          Back
        </button>
        <h1 className='text-3xl font-bold dark:text-white pt-6 text-center'>{subjectCode}</h1>
        
        {/* Add Data Section */}
        <div className='mb-4 mt-4 px-4 mx-2'>
          <h2 className='text-lg font-semibold dark:text-white mb-2 flex items-center'>
          <IoMdPersonAdd className='mr-2 text-violet-600'/>
            <span className='bg-gradient-to-r from-violet-600 to-indigo-600 text-transparent bg-clip-text'>Add Data</span>
          </h2>
          <div className='flex justify-start gap-4'>
            <button 
              className='w-48 px-4 py-2 text-white border-2 border-neutral-200 dark:border-neutral-700 rounded-lg bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-indigo-500 hover:to-violet-500 transition-all duration-300 shadow-md hover:shadow-indigo-500/20'
              onClick={() => setSchema(true)}
            >
              Define Exam Schema
            </button>
            <button
              className='w-48 px-4 py-2 text-white border-2 border-neutral-200 dark:border-neutral-700 rounded-lg bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-indigo-500 hover:to-violet-500 transition-all duration-300 shadow-md hover:shadow-indigo-500/20'
              onClick={() => setCreate(true)}
            >
              Add Student
            </button>
          </div>
        </div>

        {/* Excel Sheets Section */}
        <div className='px-4 pt-4 mx-2 w-11/12'>
          <h2 className='text-lg font-semibold dark:text-white mb-2 flex items-center'>
          <IoStatsChart className='mr-2 text-violet-600'/>
            <span className='bg-gradient-to-r from-violet-600 to-indigo-600 text-transparent bg-clip-text'>Excel Sheets</span>
          </h2>
          <div className='grid lg:grid-cols-6 md:grid-cols-3 grid-cols-2 gap-4'>
            <button className='px-4 py-2 text-white border-2 border-neutral-200 dark:border-neutral-700 rounded-lg bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-indigo-500 hover:to-violet-500 transition-all duration-300 shadow-md hover:shadow-indigo-500/20'
              onClick={() => overallSheet(subjectCode)}>
              Overall Report
            </button>  
            <button className='px-4 py-2 text-white border-2 border-neutral-200 dark:border-neutral-700 rounded-lg bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-indigo-500 hover:to-violet-500 transition-all duration-300 shadow-md hover:shadow-indigo-500/20'
              onClick={() => downloadMST1(subjectCode)}>
              MST1
            </button>  
            <button className='px-4 py-2 text-white border-2 border-neutral-200 dark:border-neutral-700 rounded-lg bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-indigo-500 hover:to-violet-500 transition-all duration-300 shadow-md hover:shadow-indigo-500/20'
              onClick={() => downloadMST2(subjectCode)}>
              MST2
            </button>  
            <button className='px-4 py-2 text-white border-2 border-neutral-200 dark:border-neutral-700 rounded-lg bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-indigo-500 hover:to-violet-500 transition-all duration-300 shadow-md hover:shadow-indigo-500/20'
              onClick={() => downloadAssignment(subjectCode)}>
              Assignment
            </button>  
            <button className='px-4 py-2 text-white border-2 border-neutral-200 dark:border-neutral-700 rounded-lg bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-indigo-500 hover:to-violet-500 transition-all duration-300 shadow-md hover:shadow-indigo-500/20'
              onClick={() => downloadEndSem(subjectCode)}>
              EndSem
            </button>  
            <button className='px-4 py-2 text-white border-2 border-neutral-200 dark:border-neutral-700 rounded-lg bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-indigo-500 hover:to-violet-500 transition-all duration-300 shadow-md hover:shadow-indigo-500/20'
              onClick={() => downloadCOSheet(subjectCode)}>
              CO Analysis
            </button>
            <button 
              className='px-4 py-2 text-white border-2 border-neutral-200 dark:border-neutral-700 rounded-lg bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-indigo-500 hover:to-violet-500 transition-all duration-300 shadow-md hover:shadow-indigo-500/20'
              onClick={() => downloadCOMatrix(subjectCode)}
            >
              CO Matrix
            </button>
          </div>
        </div>
        <List subjectCode={subjectCode}/>
      </div>
    </div>
  )
}

// Separate List component with its own state management
const List = ({ subjectCode }) => {
  // Declare all state variables at the beginning of the component
  const [sheets, setSheets] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  console.log(sheets);

  useEffect(() => {
    // Define the fetch function
    const fetchData = async () => {
      try {
        setLoading(true);
        const response = await axios.get(`http://localhost:8080/api/operation/sheets?subjectCode=${subjectCode}`);
        setSheets(response.data);
        setError(null);
        setLoading(false);
      } catch (err) {
        console.error('Error fetching sheets:', err);
        setError(err.response?.data?.error || 'Failed to fetch sheets');
      }
    };

    if (subjectCode) {
      fetchData();
    }
  }, [subjectCode]);

  if (loading) {
    return (
      <div className="m-4 p-3 bg-[#F5F5F5] dark:bg-neutral-800 rounded-lg mt-8 flex justify-center items-center">
        <div className="dark:text-white">Loading...</div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="m-4 p-3 bg-[#F5F5F5] dark:bg-neutral-800 rounded-lg mt-8">
        <div className="text-red-500 text-center">{error}</div>
      </div>
    );
  }

  return (
    <div className="m-4 md:mr-20 p-3 bg-[#F5F5F5] dark:bg-neutral-900 rounded-lg mt-8 border-[1px] dark:border-neutral-700">
      <div className="relative overflow-x-auto mx-2">
      <table className="w-full text-sm text-left rtl:text-right text-neutral-900 ">
        <thead className="text-gray-700 dark:text-white uppercase bg-grey-500 ">
          <tr className="rounded-lg">
            <th scope="col" className="px-6 py-3 text-lg">Enrollment No.</th>
            <th scope="col" className="px-6 py-3 text-lg">Name</th>

            {/* Main columns for MST1 and MST2 with sub-columns */}
            <th scope="col" colSpan="3" className="px-6 py-3 text-lg">MST1</th>
            <th scope="col" colSpan="3" className="px-6 py-3 text-lg">MST2</th>

            {/* Main column for Assignment with sub-columns */}
            <th scope="col" colSpan="5" className="px-6 py-3 text-lg">Assignment</th>
            
            {/* Main column for Endsem with sub-columns */}
            <th scope="col" colSpan="5" className="px-6 py-3 text-lg">Endsem</th>
          </tr>
          <tr className="rounded-lg">
            {/* Empty cells for Enrollment and Name */}
            <th scope="col" className="px-6 py-3"></th>
            <th scope="col" className="px-6 py-3"></th>
            
            {/* Sub-columns for MST1 */}
            <th scope="col" className="px-6 py-3">Q1</th>
            <th scope="col" className="px-6 py-3">Q2</th>
            <th scope="col" className="px-6 py-3">Q3</th>

            {/* Sub-columns for MST2 */}
            <th scope="col" className="px-6 py-3">Q1</th>
            <th scope="col" className="px-6 py-3">Q2</th>
            <th scope="col" className="px-6 py-3">Q3</th>

            {/* Sub-columns for Assignment */}
            <th scope="col" className="px-6 py-3">CO1</th>
            <th scope="col" className="px-6 py-3">CO2</th>
            <th scope="col" className="px-6 py-3">CO3</th>
            <th scope="col" className="px-6 py-3">CO4</th>
            <th scope="col" className="px-6 py-3">CO5</th>

            {/* Sub-columns for Endsem */}
            <th scope="col" className="px-6 py-3">Q1</th>
            <th scope="col" className="px-6 py-3">Q2</th>
            <th scope="col" className="px-6 py-3">Q3</th>
            <th scope="col" className="px-6 py-3">Q4</th>
            <th scope="col" className="px-6 py-3">Q5</th>
            <th scope="col" className="px-6 py-3">Remove</th>
          </tr>
        </thead>
        <tbody>
          {sheets.length === 0 ? (
            <tr className="bg-white border-[1px] dark:bg-black dark:text-gray-300">
              <td colSpan="18" className="px-6 py-4 text-center">
                No students found
              </td>
            </tr>
          ) : (
            [...sheets]
              .sort((a, b) => a.id.localeCompare(b.id))
              .map((sheet) => (
                <tr key={sheet.id} className="bg-white border-[1px] dark:border-neutral-700 dark:bg-black dark:text-gray-300">
                  <td className="px-6 py-3">{sheet.id}</td>
                  <td className="px-6 py-3">{sheet.name}</td>

                  {/* MST1 Sub-columns */}
                  <td className="px-6 py-3">{sheet.MST1_Q1 ?? '-'}</td>
                  <td className="px-6 py-3">{sheet.MST1_Q2 ?? '-'}</td>
                  <td className="px-6 py-3">{sheet.MST1_Q3 ?? '-'}</td>

                  {/* MST2 Sub-columns */}
                  <td className="px-6 py-3">{sheet.MST2_Q1 ?? '-'}</td>
                  <td className="px-6 py-3">{sheet.MST2_Q2 ?? '-'}</td>
                  <td className="px-6 py-3">{sheet.MST2_Q3 ?? '-'}</td>

                  {/* Assignment Sub-columns */}
                  <td className="px-6 py-3">{sheet.Assignment_CO1 ?? '-'}</td>
                  <td className="px-6 py-3">{sheet.Assignment_CO2 ?? '-'}</td>
                  <td className="px-6 py-3">{sheet.Assignment_CO3 ?? '-'}</td>
                  <td className="px-6 py-3">{sheet.Assignment_CO4 ?? '-'}</td>
                  <td className="px-6 py-3">{sheet.Assignment_CO5 ?? '-'}</td>

                  {/* Endsem Sub-columns */}
                  <td className="px-6 py-3">{sheet.EndSem_Q1 ?? '-'}</td>
                  <td className="px-6 py-3">{sheet.EndSem_Q2 ?? '-'}</td>
                  <td className="px-6 py-3">{sheet.EndSem_Q3 ?? '-'}</td>
                  <td className="px-6 py-3">{sheet.EndSem_Q4 ?? '-'}</td>
                  <td className="px-6 py-3">{sheet.EndSem_Q5 ?? '-'}</td>
                  <td className="px-6 py-3">
                    <button className='text-red-500' onClick={() => deleteStudent(sheet.id,subjectCode)}>
                      <FaTrash />
                    </button>
                  </td>
                </tr>
              ))
          )}
        </tbody>
      </table>
      </div>
    </div>
  );
};


const deleteStudent = async (id, subjectCode) => {
  if(window.confirm('Are you sure you want to delete this student?')){
    const response = await axios.delete(`http://localhost:8080/api/operation/sheets/${id}/${subjectCode}`);
    console.log(response.data);
    window.location.reload();
  }
}



const AddExamSchema = ({ setSchema, subjectCode }) => {
  const [formData, setFormData] = useState({
    subjectCode,
    mst1: { Q1: '', Q2: '', Q3: '' }, // Keep the structure but initialize with empty strings
    mst2: { Q1: '', Q2: '', Q3: '' }, // Keep the structure but initialize with empty strings
  });
  const [error, setError] = useState('');

  // Handle change for dropdowns
  const handleChange = (e) => {
    const { name, value } = e.target;

    // Update the corresponding mst1 or mst2 based on the input name
    if (name.startsWith('MST1_')) {
      const question = name.split('_')[1]; // Extract Q1, Q2, or Q3
      setFormData((prevData) => ({
        ...prevData,
        mst1: { ...prevData.mst1, [question]: value }, // Update the specific question
      }));
    } else if (name.startsWith('MST2_')) {
      const question = name.split('_')[1];
      setFormData((prevData) => ({
        ...prevData,
        mst2: { ...prevData.mst2, [question]: value }, // Update the specific question
      }));
    }
  };

  // Handle form submission
  const handleSubmit = async (e) => {
    e.preventDefault();

    // Prepare the data according to your specified structure
    const dataToSubmit = {
      subjectCode: formData.subjectCode,
      mst1: {
        Q1: formData.mst1.Q1,
        Q2: formData.mst1.Q2,
        Q3: formData.mst1.Q3,
      },
      mst2: {
        Q1: formData.mst2.Q1,
        Q2: formData.mst2.Q2,
        Q3: formData.mst2.Q3,
      }
    };

    console.log('Data to Submit:', dataToSubmit); // Debug log to check the final form data

    try {
      const response = await axios.post(`http://localhost:8080/api/operation/co-form`, dataToSubmit);
      alert('Form submitted successfully!');
      console.log(response.data);
      setSchema(false); // Optionally close the form
    } catch (err) {
      console.error('Error submitting form:', err);
      setError('Failed to submit the form. Please try again.');
    }
  };

  return (
    <div className="fixed h-full w-full flex items-center justify-center rounded-tl-2xl z-50 poppins-regular backdrop-brightness-50 dark:backdrop-brightness-50 backdrop-blur-sm">
      <form className="max-w-md w-1/3 mx-auto bg-white dark:bg-black rounded-xl p-2 border-2 border-neutral-300 dark:border-neutral-700" onSubmit={handleSubmit}>
        <div className="p-4">
          <div className="flex justify-end text-white cursor-pointer" onClick={() => setSchema(false)}>
            ❌
          </div>

          {/* MST1 Schema */}
          <h3 className="text-xl font-semibold text-gray-700 dark:text-gray-200 mb-2">MST 1 Schema</h3>
          <div className="mb-4 flex gap-4">
            {['Q1', 'Q2', 'Q3'].map((question, index) => (
              <div key={index} className="mb-3">
                <label htmlFor={`MST1_${question}`} className="block mb-2 text-sm font-medium text-gray-900 dark:text-white ml-2">
                  {question}
                </label>
                <select
                  id={`MST1_${question}`}
                  name={`MST1_${question}`} // Ensure name matches what you handle in handleChange
                  value={formData.mst1[question]} // Use the correct state property
                  onChange={handleChange}
                  className="w-full p-2.5 border border-gray-300 rounded-lg dark:bg-gray-700 dark:text-white dark:border-gray-600"
                >
                  <option value="">Select CO</option> {/* Default empty option */}
                  {['CO1', 'CO2', 'CO3', 'CO4', 'CO5'].map((co, idx) => (
                    <option key={idx} value={co}>{co}</option>
                  ))}
                </select>
              </div>
            ))}
          </div>

          {/* MST2 Schema */}
          <h3 className="text-xl font-semibold text-gray-700 dark:text-gray-200 mb-2">MST 2 Schema</h3>
          <div className="mb-4 flex gap-4">
            {['Q1', 'Q2', 'Q3'].map((question, index) => (
              <div key={index} className="mb-3">
                <label htmlFor={`MST2_${question}`} className="block mb-2 text-sm font-medium text-gray-900 dark:text-white ml-2">
                  {question}
                </label>
                <select
                  id={`MST2_${question}`}
                  name={`MST2_${question}`} // Ensure name matches what you handle in handleChange
                  value={formData.mst2[question]} // Use the correct state property
                  onChange={handleChange}
                  className="w-full p-2.5 border border-gray-300 rounded-lg dark:bg-gray-700 dark:text-white dark:border-gray-600"
                >
                  <option value="">Select CO</option> {/* Default empty option */}
                  {['CO1', 'CO2', 'CO3', 'CO4', 'CO5'].map((co, idx) => (
                    <option key={idx} value={co}>{co}</option>
                  ))}
                </select>
              </div>
            ))}
          </div>

          {error && <div className="text-red-500 mb-3">{error}</div>}

          <button
            type="submit"
            className="w-full px-4 py-2 text-white border-2 border-neutral-200 dark:border-neutral-700 rounded-md bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-indigo-500 hover:to-violet-500 transition-colors duration-800"
          >
            Submit
          </button>
        </div>
      </form>
    </div>
  );
};

 
const AddStudentPopup = ({ setCreate, subjectCode }) => {
  const navigate = useNavigate();
  const [formData, setFormData] = useState({
    id: '',
    subjectCode,
    name: '',
    mst1: { Q1: null, Q2: null, Q3: null },
    mst2: { Q1: null, Q2: null, Q3: null },
    assignment: { CO1: null, CO2: null, CO3: null, CO4: null, CO5: null },
    endsem: { Q1: null, Q2: null, Q3: null, Q4: null, Q5: null },
  });
  const [error, setError] = useState('');

  const handleChange = (e) => {
    const { name, value } = e.target;
  
    // Check if it's a nested field by looking for '_'
    if (name.includes('_')) {
      const [section, key] = name.split('_');
  
      setFormData((prevData) => ({
        ...prevData,
        [section.toLowerCase()]: {
          ...prevData[section.toLowerCase()],
          [key]: parseFloat(value), // Ensure the value is parsed as a float
        },
      }));
    } else {
      // Simple fields like id or name
      setFormData((prevData) => ({
        ...prevData,
        [name]: value,
      }));
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');

    try {
      const response = await axios.post(`http://localhost:8080/api/operation/submit-form`, formData, {
        headers: {
          'Authorization': `Bearer ${localStorage.getItem('token')}`,
          'Content-Type': 'application/json',
        },
      });
      console.log('Student added:', response.data);
      setCreate(false);
      navigate(0);
    } catch (err) {
      console.error('Error adding student:', err);
      setError(err.response?.data?.error || 'Failed to add student. Please try again.');
    }
  };

  return (
    <div className="absolute h-screen w-full flex items-center justify-center z-50 poppins-regular backdrop-blur-md backdrop-brightness-50">
      <form className="w-1/2 mx-auto bg-white dark:bg-black rounded-xl p-2 border-2 border-neutral-300 dark:border-neutral-700" onSubmit={handleSubmit}>
        <div className="p-4">
          <div className="flex justify-end text-white cursor-pointer" onClick={() => setCreate(false)}>
            <div>❌</div>
          </div>
          <div className='grid grid-cols-2 gap-6'>
            <div className="mb-4">
              <label htmlFor="id" className="block mb-2 dark:text-white text-lg font-semibold">Enrollment Number</label>
              <input
                type="text"
                name="id"
                id="id"
                value={formData.id}
                onChange={handleChange}
                required
                className="w-full px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
              />
            </div>

            <div className="mb-4">
              <label htmlFor="name" className="block mb-2 dark:text-white text-lg font-semibold">Student Name</label>
              <input
                type="text"
                name="name"
                id="name"
                value={formData.name}
                onChange={handleChange}
                required
                className="w-full px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
              />
            </div>

            <div>
              <div className='block mb-2 dark:text-white text-lg font-semibold'> MID SEMESTER EXAM - 1 MARKS</div>
              <div className='flex gap-3'>
                <div className="mb-4">
                  <label htmlFor="mst1_Q1" className="block mb-2 ml-2 dark:text-white">Q1</label>
                  <input
                    type="number"
                    name="mst1_Q1"
                    id="mst1_Q1"
                    value={formData.mst1.Q1}
                    onChange={handleChange}
                    className="w-full mx-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
                <div className="mb-4">
                  <label htmlFor="mst1_Q2" className="block mb-2 ml-2 dark:text-white">Q2</label>
                  <input
                    type="number"
                    name="mst1_Q2"
                    id="mst1_Q2"
                    value={formData.mst1.Q2}
                    onChange={handleChange}
                    className="w-full mx-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
                <div className="mb-4">
                  <label htmlFor="mst1_Q3" className="block mb-2 ml-2 dark:text-white">Q3</label>
                  <input
                    type="number"
                    name="mst1_Q3"
                    id="mst1_Q3"
                    value={formData.mst1.Q3}
                    onChange={handleChange}
                    className="w-full mx-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
              </div>
            </div>

            <div>
              <div className='block mb-2 dark:text-white text-lg font-semibold'> MID SEMESTER EXAM - 2 MARKS</div>
              <div className='flex gap-3'>
                <div className="mb-4">
                  <label htmlFor="mst2_Q1" className="block mb-2 ml-2 dark:text-white">Q1</label>
                  <input
                    type="number"
                    name="mst2_Q1"
                    id="mst2_Q1"
                    value={formData.mst2.Q1}
                    onChange={handleChange}
                    className="w-full mx-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
                <div className="mb-4">
                  <label htmlFor="mst2_Q2" className="block mb-2 ml-2 dark:text-white">Q2</label>
                  <input
                    type="number"
                    name="mst2_Q2"
                    id="mst2_Q2"
                    value={formData.mst2.Q2}
                    onChange={handleChange}
                    className="w-full mx-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
                <div className="mb-4">
                  <label htmlFor="mst2_Q3" className="block mb-2 ml-2 dark:text-white">Q3</label>
                  <input
                    type="number"
                    name="mst2_Q3"
                    id="mst2_Q3"
                    value={formData.mst2.Q3}
                    onChange={handleChange}
                    className="w-full mx-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
              </div>
            </div>

            <div>
              <div className='block mb-2 dark:text-white text-lg font-semibold'> ASSIGNMENT MARKS</div>
              <div className='flex gap-1'>
                <div className="mb-4">
                  <label htmlFor="assignment_CO1" className="block mb-2 dark:text-white">CO1</label>
                  <input
                    type="number"
                    name="assignment_CO1"
                    id="assignment_CO1"
                    value={formData.assignment.CO1}
                    onChange={handleChange}
                    className="w-full mr-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
                <div className="mb-4">
                  <label htmlFor="assignment_CO2" className="block mb-2 dark:text-white">CO2</label>
                  <input
                    type="number"
                    name="assignment_CO2"
                    id="assignment_CO2"
                    value={formData.assignment.CO2}
                    onChange={handleChange}
                    className="w-full mr-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
                <div className="mb-4">
                  <label htmlFor="assignment_CO3" className="block mb-2 dark:text-white">CO3</label>
                  <input
                    type="number"
                    name="assignment_CO3"
                    id="assignment_CO3"
                    value={formData.assignment.CO3}
                    onChange={handleChange}
                    className="w-full mr-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
                <div className="mb-4">
                  <label htmlFor="assignment_CO4" className="block mb-2 dark:text-white">CO4</label>
                  <input
                    type="number"
                    name="assignment_CO4"
                    id="assignment_CO4"
                    value={formData.assignment.CO4}
                    onChange={handleChange}
                    className="w-full mr-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
                <div className="mb-4">
                  <label htmlFor="assignment_CO5" className="block mb-2 dark:text-white">CO5</label>
                  <input
                    type="number"
                    name="assignment_CO5"
                    id="assignment_CO5"
                    value={formData.assignment.CO5}
                    onChange={handleChange}
                    className="w-full mr-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
              </div>
            </div>

            <div>
              <div className='block mb-2 dark:text-white text-lg font-semibold'> END SEMESTER EXAM MARKS</div>
              <div className='flex gap-1'>
                <div className="mb-4">
                  <label htmlFor="endSem_Q1" className="block mb-2 dark:text-white">Q1</label>
                  <input
                    type="number"
                    name="endSem_Q1"
                    id="endSem_Q1"
                    value={formData.endsem.Q1}
                    onChange={handleChange}
                    className="w-full mr-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
                <div className="mb-4">
                  <label htmlFor="endSem_Q2" className="block mb-2 dark:text-white">Q2</label>
                  <input
                    type="number"
                    name="endSem_Q2"
                    id="endSem_Q2"
                    value={formData.endsem.Q2}
                    onChange={handleChange}
                    className="w-full mr-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
                <div className="mb-4">
                  <label htmlFor="endSem_Q3" className="block mb-2 dark:text-white">Q3</label>
                  <input
                    type="number"
                    name="endSem_Q3"
                    id="endSem_Q3"
                    value={formData.endsem.Q3}
                    onChange={handleChange}
                    className="w-full mr-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
                <div className="mb-4">
                  <label htmlFor="endSem_Q4" className="block mb-2 dark:text-white">Q4</label>
                  <input
                    type="number"
                    name="endSem_Q4"
                    id="endSem_Q4"
                    value={formData.endsem.Q4}
                    onChange={handleChange}
                    className="w-full mr-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
                <div className="mb-4">
                  <label htmlFor="endSem_Q5" className="block mb-2 dark:text-white">Q5</label>
                  <input
                    type="number"
                    name="endSem_Q5"
                    id="endSem_Q5"
                    value={formData.endsem.Q5}
                    onChange={handleChange}
                    className="w-full mr-1 px-3 py-2 border rounded-md focus:outline-none focus:ring focus:border-blue-300 dark:bg-gray-800 dark:text-white"
                  />
                </div>
              </div>
            </div>

            {error && <div className="text-red-500 mb-3">{error}</div>}
          </div>
          <button type="submit" className="w-full px-4 py-2 mt-2 text-white border-2 border-neutral-200 dark:border-neutral-700 rounded-md bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-indigo-500 hover:to-violet-500 transition-colors duration-800">
            Submit
          </button>
        </div>
      </form>
    </div>
  );
};


  const overallSheet = (subjectCode) => {
    axios.get(`http://localhost:8080/api/operation/overall-sheet?subjectCode=${subjectCode}`, {
      responseType: 'blob', // Important to set response type as blob for file download
    })
    .then((response) => {
      // Create a URL for the blob
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Overall_Sheet.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
    })
    .catch((error) => {
      console.error('Error downloading the Excel sheet:', error);
    });
  };

  const downloadMST1 = (subjectCode) => {
    axios.get(`http://localhost:8080/api/operation/downloadmst1/${subjectCode}`, {
      responseType: 'blob', // Important to set response type as blob for file download
    })
    .then((response) => {
      // Create a URL for the blob
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const a = document.createElement('a');
      a.href = url;
      a.download = 'MST1_Sheet.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
    })
    .catch((error) => {
      console.error('Error downloading the Excel sheet:', error);
    });
  };

  const downloadMST2 = (subjectCode) => {
    axios.get(`http://localhost:8080/api/operation/downloadmst2/${subjectCode}`, {
      responseType: 'blob', // Important to set response type as blob for file download
    })
    .then((response) => {
      // Create a URL for the blob
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const a = document.createElement('a');
      a.href = url;
      a.download = 'MST2_Sheet.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
    })
    .catch((error) => {
      console.error('Error downloading the Excel sheet:', error);
    });
  };

  const downloadEndSem = (subjectCode) => {
    axios.get(`http://localhost:8080/api/operation/end-excel/${subjectCode}`, {
      responseType: 'blob',
    })
    .then((response) => {
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'EndSem_Sheet.xlsx');
      document.body.appendChild(link);
      link.click();
      link.remove();
    })
    .catch((error) => {
      console.error('Error downloading the Excel sheet:', error);
    });
  };

  const downloadAssignment = (subjectCode) => {
    axios.get(`http://localhost:8080/api/operation/assignment-excel/${subjectCode}`, {
      responseType: 'blob',
    })
    .then((response) => {
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'Assignment_Sheet.xlsx');
      document.body.appendChild(link);
      link.click();
      link.remove();
    })
    .catch((error) => {
      console.error('Error downloading the Excel sheet:', error);
    });
  };

  const downloadCOSheet = (subjectCode) => {
    axios.get(`http://localhost:8080/api/operation/generate-co-attainment/${subjectCode}`, {
      responseType: 'blob',
    })
    .then((response) => {
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'CO_Attainment_Sheet.xlsx');
      document.body.appendChild(link);
      link.click();
      link.remove();
    })
    .catch((error) => {
      console.error('Error downloading the Excel sheet:', error);
    });
  };

  const downloadCOMatrix = (subjectCode) => {
    axios.get(`http://localhost:8080/api/operation/co-matrix/${subjectCode}`, {
      responseType: 'blob',
    })
    .then((response) => {
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `CO_Matrix_${subjectCode}.xlsx`);
      document.body.appendChild(link);
      link.click();
      link.remove();
    })
    .catch((error) => {
      console.error('Error downloading the CO matrix:', error);
    });
  };

export default SubDash;