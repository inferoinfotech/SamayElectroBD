const mongoose = require('mongoose');

const TaskItemSchema = new mongoose.Schema({
  name: { type: String, required: true },
  status: { type: Boolean, default: false }
}, { _id: false });

const ClientProgressFIledSchema = new mongoose.Schema({
  clients: [{
    clientName: { type: String, required: true },
    clientId: { type: mongoose.Schema.Types.ObjectId, ref: 'MainClient', required: true, unique: true },
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

module.exports = mongoose.model('ClientProgressField', ClientProgressFIledSchema);