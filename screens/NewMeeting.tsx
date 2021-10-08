// Copyright (c) Microsoft.
// Licensed under the MIT license.

// <NewEventScreenSnippet>
import React from 'react';
import {
  ActivityIndicator,
  Alert,
  Button,
  Modal,
  Platform,
  ScrollView,
  StyleSheet,
  Text,
  View,
} from 'react-native';
import DateTimePicker from '@react-native-community/datetimepicker';
import {createStackNavigator} from '@react-navigation/stack';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import moment from 'moment-timezone';

import {UserContext} from '../UserContext';
import {GraphManager} from '../graph/GraphManager';
import {TextInput} from 'react-native-gesture-handler';

const Stack = createStackNavigator();
const NewEventState = React.createContext<NewEventState>({
  isCreating: false,
  subject: '',
  attendees: '',
  body: '',
  timeZone: '',
  startDate: new Date(),
  endDate: new Date(),
  disableCreate: () => {
    return true;
  },
  updateValue: () => {},
});

type NewEventState = {
  isCreating: boolean;
  subject: string;
  attendees: string;
  body: string;
  timeZone: string;
  startDate: Date;
  endDate: Date;
  disableCreate: () => boolean;
  updateValue: (newValue: string | Date | boolean, fieldName: string) => void;
};

type DateTimeInputProps = {
  value: Date;
  onChange: (event: Event, newValue: Date | undefined) => void;
};

// The picker acts very differently on Android and iOS
// iOS can use a single picker for both date and time,
// where Android requires two. Also the iOS version can
// be displayed all the time, while the Android version is a
// modal pop-up. Encapsulating this into a reusable component
const DateTimeInput = (props: DateTimeInputProps) => {
  const [showDatePicker, setShowDatePicker] = React.useState(false);
  const [showTimePicker, setShowTimePicker] = React.useState(
    Platform.OS === 'ios',
  );

  return (
    <View style={Platform.OS === 'android' ? styles.dateTime : null}>
      {Platform.OS === 'android' && (
        <Text
          style={styles.time}
          onPress={() => {
            setShowTimePicker(true);
          }}>
          {formatTime(props.value)}
        </Text>
      )}
      {showTimePicker && (
        <DateTimePicker
          mode={Platform.OS === 'ios' ? 'datetime' : 'time'}
          value={props.value}
          onChange={(e: Event, d: Date | undefined) => {
            setShowTimePicker(Platform.OS === 'ios');
            if (d) {
              props.onChange(e, d);
            }
          }}
        />
      )}
      {Platform.OS === 'android' && (
        <Text
          style={styles.date}
          onPress={() => {
            setShowDatePicker(true);
          }}>
          {formatDate(props.value)}
        </Text>
      )}
      {showDatePicker && Platform.OS === 'android' && (
        <DateTimePicker
          mode='date'
          value={props.value}
          onChange={(e: Event, d: Date | undefined) => {
            setShowDatePicker(Platform.OS === 'ios');
            if (d) {
              props.onChange(e, d);
            }
          }}
        />
      )}
    </View>
  );
};

const NewEventComponent = () => {
  const newEventState = React.useContext(NewEventState);
  const [meeting, setMeeting] = React.useState(null);
  const [user, setUser] = React.useState(null);

  const createMeeting = async () => {
    newEventState.updateValue(true, 'isCreating');

    const newMeeting = {
      startDateTime: moment(newEventState.startDate).format(
        'YYYY-MM-DDTHH:mm:ss',
      ),
      endDateTime: moment(newEventState.endDate).format('YYYY-MM-DDTHH:mm:ss'),
      subject: 'User Token Meeting',
    };

    const response = m => {
      console.log('receipt: ', m);
      setMeeting(m);
    };

    GraphManager.createMeeting(newMeeting, response);

    newEventState.updateValue(false, 'isCreating');
  };

  const getRecordCall = async () => {
    if (meeting) {
      const response = m => {
        console.log('receipt: ', m);
      };
      const id = meeting?.id || '';
      console.log('id: ', id);

      GraphManager.getRecordCall(id, response);
    }
  };

  const getMe = async () => {
    const u = await GraphManager.getMe();
    if (u) {
      GraphManager.createCall(u.id);
    }
  };

  const createCall = () => {
    console.log('user create: ', user);

    if (user) {
      GraphManager.createCall(user.id);
    }
  };

  return (
    <ScrollView style={styles.container}>
      <Modal visible={newEventState.isCreating}>
        <View style={styles.loading}>
          <ActivityIndicator
            color={Platform.OS === 'android' ? '#276b80' : undefined}
            animating={newEventState.isCreating}
            size='large'
          />
        </View>
      </Modal>
      <View style={styles.formField}>
        <Text style={styles.fieldLabel}>Subject</Text>
        <TextInput
          style={styles.textInput}
          value={newEventState.subject}
          onChangeText={text => newEventState.updateValue(text, 'subject')}
        />
      </View>
      <View style={styles.formField}>
        <Text style={styles.fieldLabel}>Attendees</Text>
        <TextInput
          style={styles.textInput}
          placeholder="Email (separate multiple with ';')"
          value={newEventState.attendees}
          onChangeText={text => newEventState.updateValue(text, 'attendees')}
        />
      </View>
      <View style={styles.formField}>
        <Text style={styles.fieldLabel}>Start</Text>
        <DateTimeInput
          value={newEventState.startDate}
          onChange={(e, date) => newEventState.updateValue(date!, 'startDate')}
        />
      </View>
      <View style={styles.formField}>
        <Text style={styles.fieldLabel}>End</Text>
        <DateTimeInput
          value={newEventState.endDate}
          onChange={(e, date) => newEventState.updateValue(date!, 'endDate')}
        />
      </View>
      <View style={styles.formField}>
        <TextInput
          style={styles.multiLineTextInput}
          multiline={true}
          textAlignVertical='top'
          placeholder='Body'
          value={newEventState.body}
          onChangeText={text => newEventState.updateValue(text, 'body')}
        />
      </View>
      <View style={styles.formField}>
        <Button
          title='Get me'
          disabled={newEventState.disableCreate()}
          onPress={getMe}
        />
      </View>
      <View style={styles.formField}>
        <Button
          title='create call'
          disabled={newEventState.disableCreate()}
          onPress={createCall}
        />
      </View>
      <View style={styles.formField}>
        <Button
          title='Create meeting'
          disabled={newEventState.disableCreate()}
          onPress={createMeeting}
        />
      </View>
      <View style={styles.formField}>
        <Button
          title='Get Record Call'
          disabled={newEventState.disableCreate()}
          onPress={getRecordCall}
        />
      </View>
    </ScrollView>
  );
};

const formatTime = (dateTime: Date): string => {
  return moment(dateTime).format('h:mm A');
};

const formatDate = (dateTime: Date): string => {
  return moment(dateTime).format('MMM D, YYYY');
};

// When first loading the form, set the start time
// to the nearest hour or half-hour
const getDefaultStart = (): Date => {
  const now = moment().startOf('minute');
  const offset = 30 - (now.minute() % 30);

  return now.add(offset, 'minutes').toDate();
};

// When first loading the form, set the end time
// to start + 30 min
const getDefaultEnd = (): Date => {
  return moment(getDefaultStart()).add(30, 'minutes').toDate();
};

export default class NewMeeting extends React.Component {
  static contextType = UserContext;

  // Disable the create button if:
  // - App is waiting for the result of create request
  // - Subject is empty
  // - Start time is after end time
  disableCreate = () => {
    return (
      this.state.isCreating ||
      this.state.subject.length <= 0 ||
      moment(this.state.startDate).isAfter(this.state.endDate)
    );
  };

  onStateValueChange = (
    newValue: string | Date | boolean,
    fieldName: string,
  ) => {
    this.setState({[fieldName]: newValue});
  };

  state: NewEventState = {
    isCreating: false,
    subject: '',
    attendees: '',
    body: '',
    timeZone: this.context.userTimeZone,
    startDate: getDefaultStart(),
    endDate: getDefaultEnd(),
    disableCreate: this.disableCreate,
    updateValue: this.onStateValueChange,
  };

  render() {
    return (
      <NewEventState.Provider value={this.state}>
        <Stack.Navigator>
          <Stack.Screen
            name='NewEvent'
            component={NewEventComponent}
            options={{
              headerShown: false,
            }}
          />
        </Stack.Navigator>
      </NewEventState.Provider>
    );
  }
}

const styles = StyleSheet.create({
  container: {
    flex: 1,
  },
  loading: {
    flex: 1,
    justifyContent: 'center',
    alignItems: 'center',
  },
  formField: {
    paddingHorizontal: 10,
    paddingVertical: 5,
  },
  fieldLabel: {
    fontWeight: '700',
    marginBottom: 10,
  },
  textInput: {
    borderColor: 'gray',
    borderWidth: 1,
    height: 40,
    padding: 10,
  },
  multiLineTextInput: {
    borderColor: 'gray',
    borderWidth: 1,
    height: 200,
    padding: 10,
  },
  time: {
    padding: 10,
    backgroundColor: '#e6e6e6',
    color: '#147efb',
    marginRight: 10,
  },
  date: {
    padding: 10,
    backgroundColor: '#e6e6e6',
    color: '#147efb',
  },
  dateTime: {
    flexDirection: 'row',
  },
});
// </NewEventScreenSnippet>
