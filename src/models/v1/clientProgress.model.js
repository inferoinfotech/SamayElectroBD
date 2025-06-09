const mongoose = require('mongoose');

// Task schema for each lot
const TaskItemSchema = new mongoose.Schema({
  name: { type: String, required: true },
  status: { type: Boolean, default: false }
}, { _id: false });

// Client progress schema
const ClientProgressSchema = new mongoose.Schema({
  month: { type: Number, required: true },  // 1-12
  year: { type: Number, required: true },
  clients: [{
    clientName: { type: String, required: true },
    clientId: { type: mongoose.Schema.Types.ObjectId, ref: 'MainClient', required: true },
    stageOne: [TaskItemSchema],  // Changed from tasksOfFirstLot
    stageTwo: [TaskItemSchema],  // Changed from tasksOfSecondLot
    stageThree: [TaskItemSchema],  // New field
    stageBilling: [TaskItemSchema],  // New field
    otherTasks: [{
      name: { type: String, required: true },
      status: { type: Boolean, default: false }
    }],
    remark: { type: String, default: '' }
  }]
}, { timestamps: true });

module.exports = mongoose.model('ClientProgress', ClientProgressSchema);