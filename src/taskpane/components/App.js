import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
import { PolynomialRegressor } from '@rainij/polynomial-regression-js';
import Highcharts from "highcharts";
import HighchartsReact from "highcharts-react-official";
require("highcharts/highcharts-3d")(Highcharts);
require("highcharts/modules/heatmap")(Highcharts)
require("highcharts-coloraxis-bands")(Highcharts)
require("highcharts-contour")(Highcharts)
require("highcharts-draggable-3d")(Highcharts)

/* global console, Excel, require */

const fastGridWidth = 7;
const prettyGridWidth = 21;

const App = props => {
  const [state, setState] = React.useState({
    listItems: [],
  }
  )

  const [Xaxis, setXaxis] = React.useState('x')
  const [Zaxis, setZaxis] = React.useState('y')
  const [Yaxis, setYaxis] = React.useState('z')

  React.useEffect(() => {
    setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }, [])

  const [data, setData] = React.useState([])
  const [model, setModel] = React.useState()
  console.log("🚀 ~ file: App.js:52 ~ App ~ model:", model)
  console.log("🚀 ~ file: App.js:43 ~ App ~ data:", data)

  React.useEffect(() => {
    if (data.length) {
      const x = data.filter(({ x, y, z }) => x && y && z).map(({ x, z }) => {
        return [x, z]
      })
      const y = data.filter(({ x, y, z }) => x && y && z).map(({ y }) => {
        return [y]
      })
      const model = new PolynomialRegressor(2);
      try {
        model.fit(x, y) // Training
        setModel(model);
        console.log("")
      } catch (e) {
        console.log(e.message)
      }
    }
  }, [data])

  const click = React.useCallback(async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
        */
        // const range = context.workbook.getSelectedRange();

        // Read the range address
        // range.load("address");

        // Update the fill color
        // range.format.fill.color = "yellow";

        // await context.sync();
        // console.log(`The range address was ${range.address}.`);

        let sheet = context.workbook.worksheets.getItem("Sheet1");

        let range = sheet.getUsedRangeOrNullObject();
        range.load("values");
        await context.sync();

        setData(range.values.map(([x, y, z]) => ({ x, y, z })))
      })
    } catch (error) {
      console.error(error);
    }
  }, [])

  const chartComponent = React.useRef(null)

  const setGridWidth = React.useCallback((grid_width) => { setHighchartsOptions((state) => ({ ...state, series: [{ ...state.series?.[0], grid_width }] })) }, [])

  const [highchartsOptions, setHighchartsOptions] = React.useState(() => {
    const chart = chartComponent?.current?.chart;

    return {
      chart: {
        credits: {
          enabled: false
        },
        margin: 125,
        marginBottom: 175,
        options3d: {
          enabled: true,
          alpha: 30,
          beta: 45,
          depth: 250,
          fitToPlot: false,
          axisLabelPosition: 'auto',
          drag: {
            enabled: true,
            minBeta: Number.NEGATIVE_INFINITY,
            maxBeta: Number.POSITIVE_INFINITY,
            snap: 15,
            animateSnap: true,
            beforeDrag: function () {
              setGridWidth(fastGridWidth)
            },
            afterDrag: function () {
              setGridWidth(prettyGridWidth)
            }
          },
          frame: {
            size: 10,
            visible: 'auto',
          }
        }
      },
      title: {
        text: '',
      },
      subtitle: {
        text: '',
      },
      tooltip: {
        pointFormat: 'X: <b>{point.x:.1f}</b><br/>Y: <b>{point.z:.1f}</b><br/>Z: <b>{point.y:.1f}</b>',
      },
      yAxis: {
        title: {
          text: 'Axis-Y'
        },
        labels: {
          skew3d: true,
          position3d: 'flap',
        },
        tickPixelInterval: 30,
        minPadding: 0.05,
        maxPadding: 0.05,
      },
      xAxis: {
        title: {
          text: 'Axis-X',
        },
        labels: {
          skew3d: true,
          position3d: 'flap',
        },
        tickPixelInterval: 30,
        minPadding: 0,
        maxPadding: 0,
        min: -10,
        max: 10,
      },
      zAxis: {
        title: {
          text: 'Axis-Z',
        },
        labels: {
          skew3d: true,
          position3d: 'flap',
        },
        tickPixelInterval: 30,
        minPadding: 0,
        maxPadding: 0,
        min: -10,
        max: 10,
      },
      colorAxis: {
        stops: [
          [0.0, '#3060cf'],
          [0.5, '#fffbbc'],
          [0.9, '#c4463a']
        ],
        banding: 0.5,
        tickPositioner: function () {
          if (chart) {
            return chart.yAxis[0].tickPositions.slice()
          }
        }
      }
    }
  })
  console.log("🚀 ~ file: App.js:208 ~ const[highchartsOptions,setHighchartsOptions]=React.useState ~ highchartsOptions:", highchartsOptions)

  React.useEffect(() => {
    if (model) {
      const getBound = (bound, axis) => {
        const val = Math[bound](...data?.map(obj => obj?.[axis]))
        return val
      }
      setHighchartsOptions((state) => ({
        xAxis: {
          ...state.xAxis,
          title: {
            text: Xaxis
          },
          min: getBound("min", Xaxis),
          max: getBound("max", Xaxis)
        },
        yAxis: {
          ...state.yAxis,
          title: {
            text: Zaxis
          },
          min: getBound("min", Zaxis),
          max: getBound("max", Zaxis)
        },
        zAxis: {
          ...state.zAxis,
          title: {
            text: Yaxis
          },
          min: getBound("min", Yaxis),
          max: getBound("max", Yaxis)
        },
        series: [{
          id: 'contour-series',
          type: 'contour',
          showEdges: true,
          dataFunction: (coord) => {
            const value = model.predict([[coord.x, coord.z]])[0][0]
            console.log("🚀 ~ file: App.js:247 ~ setHighchartsOptions ~ value:", value)
            if (data.some(({ y }) => y < 0)) {
              return value
            } else {
              if (value < 0) {
                return null
              } else {
                return value
              }
            }
          },
          grid_width: prettyGridWidth,
          interpolateTooltip: true,
          contours: ["value"],
        }]
      }))
    }
  }, [setGridWidth, model, data, Xaxis, Yaxis, Zaxis])

  const { title, isOfficeInitialized } = props;

  if (!isOfficeInitialized) {

    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo={require("./../../../assets/logo-filled.png")} title={props.title} message="Welcome" />
      <HighchartsReact
        ref={chartComponent}
        highcharts={Highcharts}
        options={highchartsOptions}
      />
      <HeroList message="Discover what Office Add-ins can do for you today!" items={state.listItems}>
        <p className="ms-font-l">
          Modify the source files, then click <b>Run</b>.
        </p>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
          Run
        </DefaultButton>
      </HeroList>
    </div>
  );
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};

export default App
