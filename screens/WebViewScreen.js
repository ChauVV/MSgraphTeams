import React from 'react';
import {View, Linking} from 'react-native';
import {WebView} from 'react-native-webview';

class WebViewScreen extends React.Component {
  componentDidMount() {
    const {link} = this.props.route.params;
    Linking.openURL(link);
  }

  render() {
    const {link} = this.props.route.params;
    // console.log('link receipt: ', link);
    return (
      <View style={{flex: 1, backgroundColor: 'red'}}>
        {/* <WebView source={{uri: link}} /> */}
      </View>
    );
  }
}

export default WebViewScreen;
